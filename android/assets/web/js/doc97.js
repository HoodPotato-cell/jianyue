/* Word 97-2003 (.doc) 文本提取：FIB → CLX 分片表 → 正文 */

import { parseCfb } from './ole.js';

function u16(dv, o) { return dv.getUint16(o, true); }
function u32(dv, o) { return dv.getUint32(o, true); }

function decodePieces(wd, pieces, cpLimit, lid) {
  const chunks = [];
  const dec16 = new TextDecoder('utf-16le');
  const dec8 = new TextDecoder((lid & 0xff) === 0x04 ? 'gbk' : 'windows-1252');
  for (const p of pieces) {
    let cpStart = p.cpStart;
    let cpEnd = Math.min(p.cpEnd, cpLimit);
    if (cpEnd <= cpStart) continue;
    const len = cpEnd - cpStart;
    if (p.compressed) {
      const off = p.fc + (cpStart - p.cpStart);
      chunks.push(dec8.decode(wd.subarray(off, off + len)));
    } else {
      const off = p.fc + (cpStart - p.cpStart) * 2;
      chunks.push(dec16.decode(wd.subarray(off, off + len * 2)));
    }
  }
  return chunks.join('');
}

function cleanText(raw) {
  let out = '';
  for (let i = 0; i < raw.length; i++) {
    const c = raw.charCodeAt(i);
    if (c === 0x13 || c === 0x14 || c === 0x15) { out += raw[i]; continue; } // 域标记留给 stripFields
    if (c === 0x0d) { out += '\n'; continue; }                 // 段落
    if (c === 0x07) { out += '\t'; continue; }                 // 表格单元格
    if (c === 0x0b) { out += '\n'; continue; }                 // 手动换行
    if (c === 0x0c) { out += '\n'; continue; }                 // 分页
    if (c === 0x09) { out += '\t'; continue; }
    if (c === 0x1e) { out += '\u2011'; continue; }
    if (c === 0x1f) continue;
    if (c < 0x20 && c !== 0x0a) continue;
    out += raw[i];
  }
  return out;
}

// 去掉 0x13..0x14 的域代码，只留结果
function stripFields(s) {
  let out = '';
  const stack = [];
  let inCode = 0;
  for (const ch of s) {
    if (ch === '\u0013') { stack.push(1); inCode++; continue; }
    if (ch === '\u0014') { if (stack.length) { stack[stack.length - 1] = 0; inCode = Math.max(0, inCode - 1); } continue; }
    if (ch === '\u0015') { stack.pop(); continue; }
    if (!inCode) out += ch;
  }
  return out;
}

export function parseDoc(buf) {
  const cfb = parseCfb(buf);
  if (!cfb.has('WordDocument')) throw new Error('不是 Word 文档（缺少 WordDocument 流）');
  const wd = cfb.read('WordDocument');
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  const nFib = u16(dv, 2);
  const lid = u16(dv, 6);
  const flags = u16(dv, 0x0a);
  const ccpText = u32(dv, 0x4c);
  const tableName = (flags & 0x0200) ? '1Table' : '0Table';

  let raw = '';
  if (nFib >= 105 && cfb.has(tableName)) {
    const tbl = cfb.read(tableName);
    const fcClx = u32(dv, 0x01a2);
    const lcbClx = u32(dv, 0x01a6);
    const tdv = new DataView(tbl.buffer, tbl.byteOffset, tbl.byteLength);
    let i = fcClx;
    const end = Math.min(fcClx + lcbClx, tbl.length);
    let plc = null;
    while (i < end) {
      const t = tbl[i];
      if (t === 1) { i += 3 + u16(tdv, i + 1); continue; }
      if (t === 2) {
        const lcb = u32(tdv, i + 1);
        plc = tbl.subarray(i + 5, i + 5 + lcb);
        break;
      }
      break;
    }
    if (plc) {
      const n = (plc.length - 4) / 12;
      const pdv = new DataView(plc.buffer, plc.byteOffset, plc.byteLength);
      const pieces = [];
      for (let k = 0; k < n; k++) {
        const cpStart = pdv.getUint32(k * 4, true);
        const cpEnd = pdv.getUint32((k + 1) * 4, true);
        const pcdOff = (n + 1) * 4 + k * 8;
        const fcRaw = pdv.getUint32(pcdOff + 2, true);
        const compressed = (fcRaw & 0x40000000) !== 0;
        const fc = fcRaw & 0x3fffffff;
        pieces.push({ cpStart, cpEnd, fc: compressed ? fc / 2 : fc, compressed });
      }
      raw = decodePieces(wd, pieces, ccpText || Infinity, lid);
    }
  }

  if (!raw) {
    // 很老的 Word 6/95：fcMin..fcMac 单字节
    const fcMin = u32(dv, 0x18);
    const fcMac = u32(dv, 0x1c);
    raw = new TextDecoder((lid & 0xff) === 0x04 ? 'gbk' : 'windows-1252').decode(wd.subarray(fcMin, fcMac));
  }

  const text = stripFields(cleanText(raw)).replace(/\r/g, '\n');
  const paragraphs = text.split('\n').map((s) => s.replace(/\u00a0/g, ' ').replace(/[ \t]+$/g, ''));
  while (paragraphs.length && !paragraphs[paragraphs.length - 1].trim()) paragraphs.pop();
  return { paragraphs, chars: paragraphs.join('').length };
}

export function docHtml(doc) {
  const body = doc.paragraphs.map((p) => {
    const t = p.trim();
    if (!t) return '<p class="dw-p dw-gap"><br></p>';
    const cells = p.split('\t');
    if (cells.length > 1) {
      return '<p class="dw-p dw-row">' + cells.map((c) => `<span class="dw-cell">${esc(c)}</span>`).join('') + '</p>';
    }
    return `<p class="dw-p">${esc(p)}</p>`;
  }).join('');
  return `<div class="doc-sheet old-doc">${body}</div>`;
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
