/* Excel 97-2003 (.xls / BIFF8) 读取：工作表、共享字符串、单元格数值 */

import { parseCfb } from './ole.js';

const R_BOF = 0x0809, R_EOF = 0x000a, R_BOUNDSHEET = 0x0085, R_SST = 0x00fc, R_CONTINUE = 0x003c;
const R_LABELSST = 0x00fd, R_NUMBER = 0x0203, R_RK = 0x027e, R_MULRK = 0x00bd;
const R_FORMULA = 0x0006, R_STRING = 0x0207, R_DIMENSIONS = 0x0200, R_MERGED = 0x00e5;

function records(u8) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const out = [];
  let p = 0;
  while (p + 4 <= u8.length) {
    const type = dv.getUint16(p, true);
    const len = dv.getUint16(p + 2, true);
    if (p + 4 + len > u8.length) break;
    out.push({ type, len, off: p + 4 });
    p += 4 + len;
    if (out.length > 200000) break;
  }
  return out;
}

function decodeRk(rk) {
  const fX100 = rk & 1, fInt = rk & 2;
  let v;
  if (fInt) {
    v = rk >> 2;
    if (v & 0x20000000) v -= 0x40000000;
  } else {
    const buf = new ArrayBuffer(8);
    const dv = new DataView(buf);
    dv.setUint32(0, 0, true);
    dv.setUint32(4, rk & 0xfffffffc, true);
    v = dv.getFloat64(0, true);
  }
  return fX100 ? v / 100 : v;
}

class SstReader {
  constructor(segments) {
    this.segs = segments;
    this.si = 0;
    this.off = 0;
  }
  get done() { return this.si >= this.segs.length; }
  move() {
    while (!this.done && this.off >= this.segs[this.si].length) { this.si++; this.off = 0; }
  }
  u8() { this.move(); if (this.done) throw new Error('SST 越界'); return this.segs[this.si][this.off++]; }
  u16() { return this.u8() | (this.u8() << 8); }
  u32() { return (this.u8() | (this.u8() << 8) | (this.u8() << 16) | (this.u8() << 24)) >>> 0; }
  skip(n) {
    while (n > 0) {
      this.move();
      if (this.done) return;
      const avail = this.segs[this.si].length - this.off;
      const take = Math.min(avail, n);
      this.off += take;
      n -= take;
    }
  }
  chars(cch, highByte) {
    let s = '';
    let wide = highByte;
    for (let i = 0; i < cch; i++) {
      if (!this.done && this.off >= this.segs[this.si].length) {
        // 跨到 CONTINUE：首字节是后续字符的编码标志
        this.si++; this.off = 0;
        this.move();
        if (this.done) break;
        wide = (this.segs[this.si][this.off++] & 1) !== 0;
      }
      if (this.done) break;
      const seg = this.segs[this.si];
      if (wide) { s += String.fromCharCode(seg[this.off] | (seg[this.off + 1] << 8)); this.off += 2; }
      else { s += String.fromCharCode(seg[this.off]); this.off += 1; }
    }
    return s;
  }
}

function parseSst(segments) {
  const r = new SstReader(segments);
  r.u32(); // total
  const unique = r.u32();
  const out = [];
  for (let i = 0; i < unique && !r.done; i++) {
    const cch = r.u16();
    const grbit = r.u8();
    const high = grbit & 0x01;
    const ext = grbit & 0x08;
    const rich = grbit & 0x04;
    let cRun = 0, cbExt = 0;
    if (rich) cRun = r.u16();
    if (ext) cbExt = r.u32();
    out.push(r.chars(cch, high));
    if (cRun) r.skip(cRun * 4);
    if (cbExt) r.skip(cbExt);
  }
  return out;
}

function decodeShortString(u8, off) {
  const cch = u8[off];
  const grbit = u8[off + 1];
  const wide = grbit & 1;
  let s = '';
  for (let i = 0; i < cch; i++) {
    if (wide) s += String.fromCharCode(u8[off + 2 + i * 2] | (u8[off + 3 + i * 2] << 8));
    else s += String.fromCharCode(u8[off + 2 + i]);
  }
  return s;
}

export function parseXls(buf) {
  const cfb = parseCfb(buf);
  const name = cfb.has('Workbook') ? 'Workbook' : cfb.has('Book') ? 'Book' : null;
  if (!name) throw new Error('不是 Excel 工作簿（缺少 Workbook 流）');
  const wb = cfb.read(name);
  const recs = records(wb);
  const dv = new DataView(wb.buffer, wb.byteOffset, wb.byteLength);

  const sheets = [];
  let sst = [];
  let i = 0;
  // 全局子流
  while (i < recs.length) {
    const r = recs[i];
    if (r.type === R_EOF) break;
    if (r.type === R_BOUNDSHEET) {
      const pos = dv.getUint32(r.off, true);
      const dt = wb[r.off + 5];
      const nm = decodeShortString(wb, r.off + 6);
      if (dt === 0) sheets.push({ name: nm, pos, cells: new Map(), merges: [], maxR: 0, maxC: 0 });
    } else if (r.type === R_SST) {
      const segs = [wb.subarray(r.off, r.off + r.len)];
      let j = i + 1;
      while (j < recs.length && recs[j].type === R_CONTINUE) {
        segs.push(wb.subarray(recs[j].off, recs[j].off + recs[j].len));
        j++;
      }
      try { sst = parseSst(segs); } catch (e) { sst = []; }
      i = j;
      continue;
    }
    i++;
  }

  // 逐工作表
  for (const sh of sheets) {
    let p = sh.pos;
    const endPos = Math.min(wb.length, sh.pos + 12 * 1024 * 1024);
    let depth = 0;
    let last = null;
    while (p + 4 <= endPos) {
      const type = dv.getUint16(p, true);
      const len = dv.getUint16(p + 2, true);
      const off = p + 4;
      if (off + len > wb.length) break;
      if (type === R_BOF) {
        const dt = dv.getUint16(off + 2, true);
        depth++;
        if (depth === 1 && dt !== 0x0010) { /* 可能是图表等，跳过 */ }
      } else if (type === R_EOF) {
        depth--;
        if (depth <= 0) break;
      } else if (depth <= 1) {
        if (type === R_LABELSST) {
          const row = dv.getUint16(off, true), col = dv.getUint16(off + 2, true);
          const idx = dv.getUint32(off + 6, true);
          put(sh, row, col, sst[idx] != null ? sst[idx] : '');
        } else if (type === R_NUMBER) {
          const row = dv.getUint16(off, true), col = dv.getUint16(off + 2, true);
          put(sh, row, col, dv.getFloat64(off + 6, true));
        } else if (type === R_RK) {
          const row = dv.getUint16(off, true), col = dv.getUint16(off + 2, true);
          put(sh, row, col, decodeRk(dv.getUint32(off + 6, true)));
        } else if (type === R_MULRK) {
          const row = dv.getUint16(off, true), colFirst = dv.getUint16(off + 2, true);
          const n = Math.floor((len - 6) / 6);
          for (let k = 0; k < n; k++) put(sh, row, colFirst + k, decodeRk(dv.getUint32(off + 4 + k * 6 + 2, true)));
        } else if (type === R_FORMULA) {
          const row = dv.getUint16(off, true), col = dv.getUint16(off + 2, true);
          const b6 = wb[off + 12], b7 = wb[off + 13];
          if (b6 === 0xff && b7 === 0xff) {
            const t = wb[off + 14];
            if (t === 0x00) last = { row, col };          // 结果在随后的 STRING 记录
            else if (t === 0x01) put(sh, row, col, dv.getUint16(off + 15, true) ? 'TRUE' : 'FALSE');
            else if (t === 0x03) put(sh, row, col, '');
          } else {
            put(sh, row, col, dv.getFloat64(off + 6, true));
            last = null;
          }
        } else if (type === R_STRING && last) {
          const cch = dv.getUint16(off, true);
          const wide = wb[off + 2] & 1;
          let s = '';
          for (let k = 0; k < cch; k++) {
            s += wide
              ? String.fromCharCode(wb[off + 3 + k * 2] | (wb[off + 4 + k * 2] << 8))
              : String.fromCharCode(wb[off + 3 + k]);
          }
          put(sh, last.row, last.col, s);
          last = null;
        } else if (type === R_MERGED) {
          const n = dv.getUint16(off, true);
          for (let k = 0; k < n; k++) {
            const o = off + 2 + k * 8;
            sh.merges.push({
              r1: dv.getUint16(o, true), r2: dv.getUint16(o + 2, true),
              c1: dv.getUint16(o + 4, true), c2: dv.getUint16(o + 6, true),
            });
          }
        }
      }
      p = off + len;
    }
  }

  return {
    sheets: sheets.map((s) => ({
      name: s.name,
      rows: s.maxR + 1,
      cols: s.maxC + 1,
      w: Math.min(1500, Math.max(360, (s.maxC + 1) * 76)),
      html: sheetHtml(s),
    })),
  };
}

function put(sh, row, col, v) {
  if (row > 65535 || col > 255) return;
  sh.cells.set(row + ':' + col, v);
  if (row > sh.maxR) sh.maxR = row;
  if (col > sh.maxC) sh.maxC = col;
}

function fmtNum(v) {
  if (typeof v !== 'number') return String(v == null ? '' : v);
  if (!isFinite(v)) return '';
  if (Number.isInteger(v)) return String(v);
  const s = Math.abs(v) < 1e-4 || Math.abs(v) >= 1e12 ? v.toExponential(4) : v.toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function numToCol(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function sheetHtml(sh) {
  if (!sh.cells.size) return '<div class="xls-empty">工作表为空</div>';
  const cover = new Set();
  const span = new Map();
  for (const m of sh.merges) {
    if (m.r1 === m.r2 && m.c1 === m.c2) continue;
    span.set(m.r1 + ':' + m.c1, { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
    for (let r = m.r1; r <= m.r2; r++) {
      for (let c = m.c1; c <= m.c2; c++) {
        if (r !== m.r1 || c !== m.c1) cover.add(r + ':' + c);
      }
    }
  }
  const cols = sh.maxC + 1;
  const totalW = 44 + cols * 88;
  const out = [`<div class="xls-scroll"><table class="xls old-xls" style="width:${totalW}px">`];
  out.push('<colgroup><col style="width:44px">' + '<col style="width:88px">'.repeat(cols) + '</colgroup>');
  out.push('<thead><tr><th class="xls-corner"></th>' + Array.from({ length: cols }, (_, i) => `<th class="xls-ch">${numToCol(i + 1)}</th>`).join('') + '</tr></thead><tbody>');
  for (let r = 0; r <= sh.maxR; r++) {
    out.push(`<tr><td class="xls-rh">${r + 1}</td>`);
    for (let c = 0; c <= sh.maxC; c++) {
      const k = r + ':' + c;
      if (cover.has(k)) continue;
      const v = sh.cells.get(k);
      const sp = span.get(k);
      const attr = sp ? ` rowspan="${sp.rs}" colspan="${sp.cs}"` : '';
      const cls = v == null || v === '' ? ' class="xls-blank"' : '';
      out.push(`<td${attr}${cls}>${esc(fmtNum(v))}</td>`);
    }
    out.push('</tr>');
  }
  out.push('</tbody></table></div>');
  return out.join('');
}
