/* PowerPoint 97-2003 (.ppt) 版式还原：按形状坐标绝对定位文字 */

import { parseCfb } from './ole.js';

const T_SLIDE_SIZE = 0x03e9;
const T_SLIDE = 0x03ee;
const T_SP_CONTAINER = 0x0f004;
const T_CLIENT_ANCHOR = 0x0f010;
const T_TEXT_BOX = 0x0f00d;
const T_TEXT_HEADER = 0x0f9f;
const T_TEXT_CHARS = 0x0fa0;
const T_TEXT_BYTES = 0x0fa8;
const T_SLIDE_LIST = 0x0ff0;

function decodeBytes(bytes) {
  let ascii = true;
  for (const b of bytes) if (b >= 0x80) { ascii = false; break; }
  if (ascii) {
    let s = '';
    for (const b of bytes) s += String.fromCharCode(b);
    return s;
  }
  try {
    const s = new TextDecoder('gbk').decode(bytes);
    if (!s.includes('\ufffd')) return s;
  } catch (e) { /* 忽略 */ }
  return new TextDecoder('windows-1252').decode(bytes);
}

function walkAll(u8, cb) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let p = 0;
  while (p + 8 <= u8.length) {
    const vi = dv.getUint16(p, true);
    const type = dv.getUint16(p + 2, true);
    const len = dv.getUint32(p + 4, true);
    if (p + 8 + len > u8.length) return;
    const body = u8.subarray(p + 8, p + 8 + len);
    if (cb(type, vi & 0xf, vi >> 4, body) === false) return;
    if ((vi & 0xf) === 0xf && len > 0) walkAll(body, cb);
    p += 8 + len;
  }
}

function eachRecord(u8, cb) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let p = 0;
  while (p + 8 <= u8.length) {
    const vi = dv.getUint16(p, true);
    const type = dv.getUint16(p + 2, true);
    const len = dv.getUint32(p + 4, true);
    if (p + 8 + len > u8.length) return;
    if (cb(type, vi & 0xf, u8.subarray(p + 8, p + 8 + len)) === false) return;
    p += 8 + len;
  }
}

function findImage(bytes) {
  const n = Math.min(bytes.length - 4, 72);
  for (let i = 0; i < n; i++) {
    if (bytes[i] === 0xff && bytes[i + 1] === 0xd8 && bytes[i + 2] === 0xff) return { mime: 'image/jpeg', start: i };
    if (bytes[i] === 0x89 && bytes[i + 1] === 0x50 && bytes[i + 2] === 0x4e && bytes[i + 3] === 0x47) return { mime: 'image/png', start: i };
  }
  return null;
}

function toDataUrl(bytes, mime) {
  let s = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    s += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return 'data:' + mime + ';base64,' + btoa(s);
}

function blipIndex(bytes) {
  for (let i = 0; i + 6 <= bytes.length; i++) {
    if (bytes[i] === 0x04 && bytes[i + 1] === 0x41) {
      const idx = bytes[i + 2] | (bytes[i + 3] << 8) | (bytes[i + 4] << 16) | (bytes[i + 5] << 24);
      if (idx > 0 && idx < 200000) return idx;
    }
  }
  return 0;
}

function cleanLines(s) {
  return String(s)
    .replace(/\r\n?/g, '\n')
    .replace(/\x0b/g, '\n')
    .replace(/[\x00-\x08\x0e-\x1f]/g, '')
    .split('\n')
    .map((x) => x.replace(/\s+$/g, '').replace(/^\s+/, ''))
    .filter(Boolean);
}

function readBlips(cfb) {
  const out = [];
  if (!cfb.has('Pictures')) return out;
  const pic = cfb.read('Pictures');
  const cache = [];
  eachRecord(pic, (type, ver, b) => {
    const img = (type >= 0xf018 && type <= 0xf01e) ? findImage(b) : null;
    if (img) {
      const bytes = b.subarray(img.start);
      out.push({ mime: img.mime, bytes });
    } else {
      out.push(null);
    }
    return true;
  });
  out.dataUrl = (i) => {
    if (!out[i]) return '';
    if (!cache[i]) cache[i] = toDataUrl(out[i].bytes, out[i].mime);
    return cache[i];
  };
  return out;
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* 一个 SpContainer 里的锚点（1/8 磅）与文字 */

/* 配色表：优先用文件里的 ColorSchemeAtom（8 色），否则退回 Office 默认 */
let SCHEME = [0xffffff, 0x000000, 0x808080, 0x000000, 0x999999, 0xcccccc, 0xffffcc, 0xb3b300];

function readColorScheme(stream) {
  let out = null;
  walkAll(stream, (type, ver, inst, body) => {
    if (type === 0x07f0 && body.length >= 32) {
      const dv = new DataView(body.buffer, body.byteOffset, body.byteLength);
      const arr = [];
      for (let i = 0; i < 8; i++) {
        const v = dv.getUint32(i * 4, true) & 0xffffff;
        const r = v & 0xff, g = (v >> 8) & 0xff, b = (v >> 16) & 0xff;
        arr.push((r << 16) | (g << 8) | b);
      }
      out = arr;
      return false;
    }
    return true;
  });
  return out;
}

function parseOpt(bytes) {
  const out = new Map();
  let i = 0;
  while (i + 6 <= bytes.length) {
    const raw = bytes[i] | (bytes[i + 1] << 8);
    i += 2;
    const num = raw & 0x3fff;
    if (raw & 0x8000) break; // 复杂属性（图形路径/数组）暂不解析
    const val = (bytes[i] | (bytes[i + 1] << 8) | (bytes[i + 2] << 16) | (bytes[i + 3] << 24)) >>> 0;
    i += 4;
    out.set(num, val);
  }
  return out;
}

function colorOf(v) {
  if (v == null) return null;
  if ((v & 0xffffff) === 0xffffff) return null;
  if ((v & 0x00ffffff) === 0 && (v >>> 24) >= 1 && (v >>> 24) <= 8) {
    const rgb2 = SCHEME[(v >>> 24)] != null ? SCHEME[(v >>> 24)] : 0;
    return '#' + rgb2.toString(16).padStart(6, '0');
  }
  if ((v & 0xff000000) === 0x08000000) {
    const idx = v & 0xff;
    const rgb = SCHEME[idx] != null ? SCHEME[idx] : 0x000000;
    return '#' + rgb.toString(16).padStart(6, '0');
  }
  const r = v & 0xff, g = (v >> 8) & 0xff, b = (v >> 16) & 0xff;
  return '#' + ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0');
}


/* ---- 段落/字符样式（StyleTextPropAtom，按 MS-PPT 2.9.20） ---- */
const PARA_PROPS = [
  { m: 0x0000000f, s: 2, k: 'flags' },
  { m: 0x00000080, s: 2, k: 'bulletChar' },
  { m: 0x00000010, s: 2, k: 'bulletFont' },
  { m: 0x00000040, s: 2, k: 'bulletSize' },
  { m: 0x00000020, s: 4, k: 'bulletColor' },
  { m: 0x00000800, s: 2, k: 'align' },
  { m: 0x00001000, s: 2, k: 'lineSpacing' },
  { m: 0x00002000, s: 2, k: 'spaceBefore' },
  { m: 0x00004000, s: 2, k: 'spaceAfter' },
  { m: 0x00000100, s: 2, k: 'textOffset' },
  { m: 0x00000400, s: 2, k: 'bulletOffset' },
  { m: 0x00008000, s: 2, k: 'defaultTab' },
  { m: 0x00100000, s: -1, k: 'tabStops' },
  { m: 0x00010000, s: 2, k: 'fontAlign' },
  { m: 0x000e0000, s: 2, k: 'wrapFlags' },
  { m: 0x00200000, s: 2, k: 'textDir' },
  { m: 0x00800000, s: 0 }, { m: 0x01000000, s: 0 }, { m: 0x02000000, s: 0 },
];
const CHAR_PROPS = [
  { m: 0x00100000, s: 0 }, { m: 0x01000000, s: 0 }, { m: 0x02000000, s: 0 }, { m: 0x04000000, s: 0 },
  { m: 0x0000ffff, s: 2, k: 'flags' },
  { m: 0x00010000, s: 2, k: 'fontIdx' },
  { m: 0x00200000, s: 2, k: 'asianFontIdx' },
  { m: 0x00400000, s: 2, k: 'ansiFontIdx' },
  { m: 0x00800000, s: 2, k: 'symbolFontIdx' },
  { m: 0x00020000, s: 2, k: 'size' },
  { m: 0x00040000, s: 4, k: 'color' },
  { m: 0x00080000, s: 2, k: 'superScript' },
];

function readStyleBlock(mask, table, bytes, pos) {
  const props = {};
  let used = 0;
  for (const p of table) {
    if ((mask & p.m) === 0) continue;
    if (p.s === 0) continue;
    if (p.s === -1) {
      if (pos + used + 2 > bytes.length) return { props, used };
      const n = bytes[pos + used] | (bytes[pos + used + 1] << 8);
      used += 2 + n * 4;
      continue;
    }
    if (pos + used + p.s > bytes.length) return { props, used };
    const off = pos + used;
    let v;
    if (p.s === 2) v = bytes[off] | (bytes[off + 1] << 8);
    else v = (bytes[off] | (bytes[off + 1] << 8) | (bytes[off + 2] << 16) | (bytes[off + 3] << 24)) >>> 0;
    used += p.s;
    if (p.k) props[p.k] = v;
  }
  return { props, used };
}

function parseStyleAtom(bytes, textLen) {
  const out = { paras: [], chars: [] };
  let pos = 0;
  let handled = 0;
  let guard = 0;
  while (pos + 10 <= bytes.length && handled < textLen && guard++ < 400) {
    const len = (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16) | (bytes[pos + 3] << 24)) >>> 0;
    pos += 4;
    const indent = bytes[pos] | (bytes[pos + 1] << 8);
    pos += 2;
    const mask = (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16) | (bytes[pos + 3] << 24)) >>> 0;
    pos += 4;
    const r = readStyleBlock(mask, PARA_PROPS, bytes, pos);
    pos += r.used;
    out.paras.push(Object.assign({ len, indent }, r.props));
    handled += len;
    if (handled >= textLen && pos + 10 <= bytes.length) { /* 允许一个额外段落 */ }
  }
  handled = 0;
  guard = 0;
  while (pos + 8 <= bytes.length && handled < textLen && guard++ < 800) {
    const len = (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16) | (bytes[pos + 3] << 24)) >>> 0;
    pos += 4;
    const mask = (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16) | (bytes[pos + 3] << 24)) >>> 0;
    pos += 4;
    const r = readStyleBlock(mask, CHAR_PROPS, bytes, pos);
    pos += r.used;
    out.chars.push(Object.assign({ len }, r.props));
    handled += len;
  }
  return out;
}

function styleColor(v) {
  if (v == null) return null;
  if ((v & 0x00ffffff) === 0 && (v >>> 24) >= 1 && (v >>> 24) <= 8) {
    const rgb3 = SCHEME[(v >>> 24)] != null ? SCHEME[(v >>> 24)] : 0;
    return '#' + rgb3.toString(16).padStart(6, '0');
  }
  if ((v & 0xff000000) === 0x08000000) {
    const idx = v & 0xff;
    const rgb = SCHEME[idx] != null ? SCHEME[idx] : 0;
    return '#' + rgb.toString(16).padStart(6, '0');
  }
  if ((v & 0xffffff) === 0xffffff) return null;
  const r = v & 0xff, g = (v >> 8) & 0xff, b = (v >> 16) & 0xff;
  return '#' + ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0');
}


/* 母版样式表：按文本类型（标题/正文）给默认字号、符号 */
function parseMasterStyle(bytes, textType) {
  if (bytes.length < 2) return null;
  const levels = bytes[0] | (bytes[1] << 8);
  let pos = 2;
  const paras = [];
  const chars = [];
  for (let i = 0; i < levels; i++) {
    if (textType >= 6) {
      if (pos + 2 > bytes.length) break;
      pos += 2; // indentLevel
    }
    if (pos + 4 > bytes.length) break;
    const pMask = (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16) | (bytes[pos + 3] << 24)) >>> 0;
    pos += 4;
    const pr = readStyleBlock(pMask, PARA_PROPS, bytes, pos);
    pos += pr.used;
    paras.push(pr.props);
    if (pos + 4 > bytes.length) break;
    const cMask = (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16) | (bytes[pos + 3] << 24)) >>> 0;
    pos += 4;
    const cr = readStyleBlock(cMask, CHAR_PROPS, bytes, pos);
    pos += cr.used;
    chars.push(cr.props);
  }
  return { paras, chars };
}

function readMasterStyles(stream) {
  const map = new Map();
  walkAll(stream, (type, ver, inst, body) => {
    if (type === 0x0fa3) {
      const t = parseMasterStyle(body, inst);
      if (t && t.chars.length) map.set(inst, t);
    }
    return true;
  });
  return map;
}

function readShape(body) {
  let anchor = null;
  let pic = 0;
  let spType = 0;
  let opt = null;
  let styleAtom = null;
  let textLen = 0;
  let hdrType = -1;
  let pptType = 0;
  let isOle = false;
  const lines = [];
  let hasText = false;
  const walk = (u8) => {
    eachRecord(u8, (type, ver, b) => {
      if (type === T_CLIENT_ANCHOR && b.length >= 8) {
        if (!anchor) {
          const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
          anchor = [dv.getUint16(0, true), dv.getUint16(2, true), dv.getUint16(4, true), dv.getUint16(6, true)];
        }
      } else if (type === 0xf00a && b.length >= 8) {
        const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
        spType = (dv.getUint32(4, true) >>> 16) & 0xffff;
      } else if (type === 0xf00b) {
        if (!pic) pic = blipIndex(b);
        if (!opt) opt = parseOpt(b);
      } else if (type === T_TEXT_CHARS) {
        hasText = true;
        const raw = new TextDecoder('utf-16le').decode(b);
        textLen += raw.length;
        for (const l of cleanLines(raw)) lines.push(l);
      } else if (type === T_TEXT_BYTES) {
        hasText = true;
        const raw = decodeBytes(b);
        textLen += raw.length;
        for (const l of cleanLines(raw)) lines.push(l);
      }
      if (type === 0x0f9f && b.length >= 4) hdrType = b[0] | (b[1] << 8);
      if (type === 0x0bc1 && b.length >= 4) pptType = b[0] | (b[1] << 8) | (b[2] << 16) | (b[3] << 24);
      if (type === 0x0fba) isOle = true;
      if (type === 0x0fa1) styleAtom = b;
      if (ver === 0xf && b.length) walk(b);
      return true;
    });
  };
  walk(body);
  const style = styleAtom ? parseStyleAtom(styleAtom, textLen) : null;
  return { anchor, lines, hasText, pic, spType, opt, style, hdrType, pptType, isOle };
}

function slideShapes(body) {
  const shapes = [];
  const walk = (u8) => {
    eachRecord(u8, (type, ver, b) => {
      if (type === T_SP_CONTAINER) {
        const s = readShape(b);
        if (s.anchor) shapes.push(s);
      }
      if (ver === 0xf && b.length) walk(b);
      return true;
    });
  };
  walk(body);
  return shapes;
}

function shapeHtml(s, pageW, blips, masters) {
  const [top, left, right, bottom] = s.anchor;
  const x = left / 8;
  const y = top / 8;
  const w = Math.max(0.5, (right - left) / 8);
  const h = Math.max(0.5, (bottom - top) / 8);
  if (s.pic && blips && !blips[s.pic - 1]) return ''; // 无法解码的图元
  if (!s.lines.length && !s.pic && (s.pptType || 0) >= 100) return ''; // OLE/流程图占位，无预览图可画 // 引用了无法解码的图元（多为 EMF/OLE 预览）
  let imgHtml = '';
  if (s.pic && blips && blips[s.pic - 1]) {
    imgHtml = `<img src="${blips.dataUrl(s.pic - 1)}" style="position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${w.toFixed(1)}px;height:${h.toFixed(1)}px;object-fit:fill" alt="">`;
  }

  // 形状填充与边框
  const opt = s.opt || new Map();
  const undecodablePic = s.pic && blips && !blips[s.pic - 1];
  let fill = undecodablePic ? null : colorOf(opt.get(0x0181));
  if (fill === '#ffffff') fill = null; // 白底多半是“无填充”，不要糊住内容
  const line = undecodablePic ? null : colorOf(opt.get(0x01c0));
  const emu = (v) => Math.round((v / 12700) * (96 / 72) * 10) / 10;
  const lineW = opt.has(0x01cb) ? Math.max(1, emu(opt.get(0x01cb))) : 1;
  const padL = opt.has(0x0081) ? emu(opt.get(0x0081)) : 4;
  const padT = opt.has(0x0082) ? emu(opt.get(0x0082)) : 2;
  const padR = opt.has(0x0083) ? emu(opt.get(0x0083)) : 4;
  const padB = opt.has(0x0084) ? emu(opt.get(0x0084)) : 2;
  const anchorV = opt.has(0x0087) ? (opt.get(0x0087) & 0xffff) : 0; // 0 上 1 中 2 下

  // 线条：类型 20；宽或高很小
  if (s.spType === 20 || (s.spType === 0 && (h < 3 || w < 3))) {
    const x1 = x, y1 = y, x2 = x + w, y2 = y + h;
    const len = Math.hypot(x2 - x1, y2 - y1);
    const ang = Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI;
    return `<div style="position:absolute;left:${x1.toFixed(1)}px;top:${y1.toFixed(1)}px;width:${Math.max(1, len).toFixed(1)}px;height:${Math.max(1, lineW).toFixed(1)}px;`
      + `background:${line || '#333'};transform-origin:0 50%;transform:rotate(${ang.toFixed(2)}deg)"></div>`;
  }
  const shapeCss = [];
  if (fill) shapeCss.push('background:' + fill);
  if (line) shapeCss.push('border:' + lineW + 'px solid ' + line);
  const pt = s.pptType || 0;
  if ((h <= 6 || w <= 6) && (fill || line)) {
    const col = fill || line;
    if (h <= 6) {
      return `<div style="position:absolute;left:${x.toFixed(1)}px;top:${(y + h / 2).toFixed(1)}px;width:${w.toFixed(1)}px;height:1px;background:${col}"></div>`;
    }
    return `<div style="position:absolute;left:${(x + w / 2).toFixed(1)}px;top:${y.toFixed(1)}px;width:1px;height:${h.toFixed(1)}px;background:${col}"></div>`;
  }
  if (pt === 5 || s.spType === 2) shapeCss.push('border-radius:10%');
  if (pt === 9 || pt === 10 || s.spType === 3) shapeCss.push('border-radius:50%');
  const shapeStyle = shapeCss.length
    ? `position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${w.toFixed(1)}px;height:${h.toFixed(1)}px;box-sizing:border-box;${shapeCss.join(';')}`
    : '';

  if (!s.lines.length) {
    return (shapeStyle ? `<div style="${shapeStyle}"></div>` : '') + imgHtml;
  }
  const mst = masters && masters.get(s.hdrType >= 0 ? s.hdrType : 1);
  const mChar = mst && (mst.chars[0] || {});
  const mPara = mst && (mst.paras[0] || {});
  const st0 = s.style ? (s.style.chars.find((c) => c.size) || s.style.chars[0] || {}) : {};
  const pa0 = s.style ? (s.style.paras.find((pp) => pp.bulletChar) || s.style.paras[0] || {}) : {};
  if (!st0.size && mChar && mChar.size) st0.size = mChar.size;
  if (st0.flags == null && mChar && mChar.flags != null) st0.flags = mChar.flags;
  if (st0.color == null && mChar && mChar.color != null) st0.color = mChar.color;
  const isTitle = s.hdrType === 0 || (s.lines.length === 1 && h >= 34 && w >= 150);
  const bulletCh = isTitle ? 0 : (pa0.bulletChar || (mPara && mPara.bulletChar) || (s.hdrType === 1 && s.lines.length > 1 ? 0x2022 : 0));
  const bulletStr = bulletCh ? (bulletCh === 0x6e ? '▪ ' : bulletCh === 0xf0b7 ? '• ' : String.fromCharCode(bulletCh) + ' ') : '';
  const linesText = bulletStr ? s.lines.map((l) => bulletStr + l) : s.lines;
  const text = linesText.join('\n');
  let size = st0.size ? Math.max(6, Math.min(96, st0.size)) : 0; // 画布 720x540 = 720x540pt，1pt = 1px
  if (!size) {
    size = Math.min(34, Math.max(10, Math.min(h / Math.max(1, s.lines.length), w / Math.max(4, (text.length / Math.max(1, s.lines.length)) * 0.62))));
    if (s.lines.length === 1 && h > 34 && w > 200) size = Math.min(34, h * 0.58);
  }
  const bold = (st0.flags != null) ? ((st0.flags & 1) !== 0) : (s.lines.length === 1 && h > 34);
  const textColor = styleColor(st0.color) || '#1a1a1a';
  {
    const cpl = Math.max(3, (w - padL - padR - (bulletStr ? size * 1.1 : 0)) / Math.max(1, size * 0.95));
    let need = 0;
    for (const ln of s.lines) need += Math.max(1, Math.ceil(ln.length / cpl));
    const needH = need * size * 1.22;
    if (needH > h && needH > 0) size = Math.max(6, size * (h / needH));
  }
  if (isTitle && s.lines.length === 1) {
    const fit = Math.min(48, h * 0.5);
    if (fit > size) size = fit;
  }
  if (s.lines.length === 1 && !bulletStr) {
    const est = s.lines[0].length * size * 0.62;
    if (est > w * 0.94) size = Math.max(8, (w * 0.94) / (s.lines[0].length * 0.95));
  }
  const pad = Math.max(1, Math.min(6, Math.round(w * 0.012)));
  const base = shapeStyle || `position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${w.toFixed(1)}px;height:${h.toFixed(1)}px;box-sizing:border-box;`;
  const alignSelf = anchorV === 1 ? 'center' : anchorV === 2 ? 'flex-end' : 'flex-start';
  return `<div class="ppt-shape" style="${base}`
    + `display:flex;flex-direction:column;justify-content:${alignSelf};`
    + `font-size:${size.toFixed(1)}px;line-height:1.12;${bold ? 'font-weight:700;' : ''}`
    + `color:${textColor};padding:${padT}px ${padR}px ${padB}px ${padL}px;overflow:hidden;white-space:pre-wrap;word-break:break-word;`
    + (bulletStr ? `text-indent:-1.1em;padding-left:${(padL + size * 1.1).toFixed(1)}px;` : '')
    + `">`
    + `<span>${esc(text)}</span></div>` + imgHtml;
}

/* 渲染后自动缩字号，直到文字装进形状 */
export function fitPptText(root) {
  if (!root || !root.querySelectorAll) return;
  for (const el of root.querySelectorAll('.ppt-shape')) {
    const orig = parseFloat(el.style.fontSize) || 16;
    let size = orig;
    let guard = 0;
    while (guard++ < 10 && size > 8 && size > orig * 0.86
      && (el.scrollHeight > el.clientHeight + 1 || el.scrollWidth > el.clientWidth + 1)) {
      size *= 0.96;
      el.style.fontSize = size.toFixed(1) + 'px';
    }
  }
}

export function parsePpt(buf) {
  const cfb = parseCfb(buf);
  if (!cfb.has('PowerPoint Document')) throw new Error('不是 PowerPoint 文件（缺少 PowerPoint Document 流）');
  const stream = cfb.read('PowerPoint Document');
  const blips = readBlips(cfb);
  const scheme = readColorScheme(stream);
  const masters = readMasterStyles(stream);
  if (scheme) SCHEME = scheme;

  let slideW = 720;
  let slideH = 540;
  eachRecord(stream, (type, ver, body) => {
    if (type === T_SLIDE_SIZE && body.length >= 8) {
      const dv = new DataView(body.buffer, body.byteOffset, body.byteLength);
      const cx = dv.getUint32(0, true);
      const cy = dv.getUint32(4, true);
      if (cx > 0 && cy > 0) {
        slideW = Math.round(cx / 8);
        slideH = Math.round(cy / 8);
      }
      return false;
    }
    return true;
  });

  // 母版装饰：取第一个 Master 记录的图形，画在每页底层
  let masterHtml = '';
  eachRecord(stream, (type, ver, body) => {
    if (type === 0x03f8 && !masterHtml) {
      const shapes = slideShapes(body).filter((s) => {
        if (s.lines.length) return false;
        if (!(s.opt || s.pic)) return false;
        const [top, left, right, bottom] = s.anchor;
        const w2 = (right - left) / 8, h2 = (bottom - top) / 8;
        // 母版只保留“装饰性”元素：细线条或小色块；大块占位框不画（会盖住正文）
        const thin = h2 <= 6 || w2 <= 6;
        if (!thin && w2 * h2 > 20000) return false;
        return true;
      });
      masterHtml = shapes.map((s) => shapeHtml(s, 720, blips, null)).join('');
      return false;
    }
    return true;
  });

  const pages = [];
  const seen = new Set();
  eachRecord(stream, (type, ver, body) => {
    if (type !== T_SLIDE) return true;
    const shapes = slideShapes(body).filter((s) => {
      if (s.lines.length || s.pic) return true;
      const o = s.opt;
      if (!o) return false;
      if (o.has(0x0181) || o.has(0x01c0)) return true;
      const hh = (s.anchor[3] - s.anchor[0]) / 8;
      const ww = (s.anchor[2] - s.anchor[1]) / 8;
      return hh < 3 || ww < 3;
    });
    if (!shapes.length) return true;
    if (globalThis.__PPT_DEBUG) console.log('[shapes]', blips.length, blips.filter(Boolean).length, shapes.map((s) => s.pic));
    const fp = shapes.map((s) => s.anchor.join(',') + '|' + s.lines.join('\u0001') + '|' + s.pic).join('\u0002');
    if (seen.has(fp)) return true;
    seen.add(fp);
    const rank = (s) => (s.lines.length ? 2 : (s.pic ? 1 : 0));
    const ordered = shapes.slice().sort((a, b) => rank(a) - rank(b));
    const html = '<div class="ppt-canvas ppt-layout" style="position:relative;width:' + slideW + 'px;height:' + slideH + 'px;background:#fff">' + masterHtml
      + ordered.map((s) => shapeHtml(s, slideW, blips, masters)).join('')
      + '</div>';
    pages.push({ w: slideW, h: slideH, html });
    return true;
  });

  // 兜底：没有形状坐标时退回纯文字大纲
  if (!pages.length) {
    eachRecord(stream, (type, inst, body) => {
      if (type !== T_SLIDE_LIST || ((inst >> 4) !== 0)) return true;
      const lines = [];
      eachRecord(body, (t, v, b) => {
        if (t === T_TEXT_CHARS) for (const l of cleanLines(new TextDecoder('utf-16le').decode(b))) lines.push(l);
        else if (t === T_TEXT_BYTES) for (const l of cleanLines(decodeBytes(b))) lines.push(l);
        return true;
      });
      if (lines.length) {
        const html = '<div class="ppt-canvas old-ppt" style="width:' + slideW + 'px;height:' + slideH + 'px;box-sizing:border-box;padding:40px;background:#fff">'
          + '<div class="old-ppt-title">' + esc(lines[0]) + '</div>'
          + lines.slice(1).map((l) => '<div class="ppt-p">' + esc(l) + '</div>').join('')
          + '</div>';
        pages.push({ w: slideW, h: slideH, html });
      }
      return true;
    });
  }

  return { pageW: slideW, pageH: slideH, pages: pages.slice(0, 300) };
}
