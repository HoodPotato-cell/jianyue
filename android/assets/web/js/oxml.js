/* OOXML 公共工具：XML 解析、单位换算、颜色、数字格式 */

export const TWIP = 15;          // 1px = 15 twips (96dpi)
export const EMU = 9525;         // 1px = 9525 EMU
export const PT = 96 / 72;       // 1pt = 1.3333px

export function parseXml(data) {
  const s = typeof data === 'string' ? data : new TextDecoder('utf-8').decode(data);
  return new DOMParser().parseFromString(s, 'application/xml');
}

/** 子元素列表（跨浏览器/Node 兼容） */
export function kids(node) {
  const out = [];
  if (!node) return out;
  const list = node.childNodes;
  if (!list) return out;
  for (let i = 0; i < list.length; i++) {
    const c = list.item ? list.item(i) : list[i];
    if (c && c.nodeType === 1) out.push(c);
  }
  return out;
}

/** 按 localName 取所有后代元素（忽略命名空间前缀） */
export function els(node, name) {
  const out = [];
  if (!node) return out;
  const walk = (n) => {
    const list = n.childNodes;
    if (!list) return;
    for (let i = 0; i < list.length; i++) {
      const c = list.item ? list.item(i) : list[i];
      if (!c || c.nodeType !== 1) continue;
      if (!name || c.localName === name) out.push(c);
      walk(c);
    }
  };
  walk(node);
  return out;
}

export function el(node, name) {
  if (!node) return null;
  let found = null;
  const walk = (n) => {
    const list = n.childNodes;
    if (!list) return;
    for (let i = 0; i < list.length; i++) {
      const c = list.item ? list.item(i) : list[i];
      if (!c || c.nodeType !== 1) continue;
      if (c.localName === name) { found = c; return; }
      walk(c);
      if (found) return;
    }
  };
  walk(node);
  return found;
}

/** 只在直接子元素里找 */
export function child(node, name) {
  if (!node) return null;
  for (const c of kids(node)) if (!name || c.localName === name) return c;
  return null;
}

export function children(node, name) {
  const out = [];
  if (!node) return [];
  for (const c of kids(node)) if (!name || c.localName === name) out.push(c);
  return out;
}

export function attr(node, name, def = null) {
  if (!node) return def;
  const attrs = node.attributes;
  if (!attrs) return def;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item ? attrs.item(i) : attrs[i];
    if (a && (a.name === name || a.localName === name)) return a.value;
  }
  return def;
}

export function num(node, name, def = 0) {
  const v = attr(node, name);
  if (v === null || v === '') return def;
  const f = parseFloat(v);
  return isNaN(f) ? def : f;
}

export function textOf(node) {
  return node ? (node.textContent || '') : '';
}

export function onOff(node, name) {
  if (!node) return false;
  const e = child(node, name);
  if (!e) return false;
  const v = attr(e, 'val');
  return v === null || v === '1' || v === 'true' || v === 'on';
}

export function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/** "auto"/"RRGGBB"/"#RRGGBB"/"FFRRGGBB" -> css color or null */
export function color(v) {
  if (!v || v === 'auto') return null;
  v = String(v).trim();
  if (v.startsWith('#')) v = v.slice(1);
  if (v.length === 8) v = v.slice(2);
  if (v.length === 6 && /^[0-9a-fA-F]+$/.test(v)) return '#' + v.toLowerCase();
  return null;
}

export function themeColor(node) {
  // a:schemeClr val="accent1" 等 -> 近似色
  const tbl = {
    accent1: '#4472c4', accent2: '#ed7d31', accent3: '#a5a5a5', accent4: '#ffc000',
    accent5: '#5b9bd5', accent6: '#70ad47', tx1: '#000000', tx2: '#44546a',
    bg1: '#ffffff', bg2: '#e7e6e6', dk1: '#000000', lt1: '#ffffff',
    dk2: '#44546a', lt2: '#e7e6e6', hlink: '#0563c1', folHlink: '#954f72',
    phClr: '#000000', none: null,
  };
  const v = attr(node, 'val');
  return tbl[v] || null;
}

export function px(v) {
  return Math.round(v * 100) / 100;
}

/** Excel 序列号 -> Date */
export function excelSerialToDate(serial) {
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  return new Date(ms);
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function pad(n, w) { let s = String(Math.abs(n)); while (s.length < w) s = '0' + s; return n < 0 ? '-' + s : s; }

/** 简化版 Excel 数字格式应用 */
/** 把格式串切成 token：{t:'lit'|'y'|'m'|'d'|'h'|'s'|'ampm'|'num', s:原文} */
function tokenize(fmt) {
  const toks = [];
  let i = 0;
  const push = (t, s) => toks.push({ t, s });
  while (i < fmt.length) {
    const c = fmt[i];
    if (c === '"') {
      let j = fmt.indexOf('"', i + 1);
      if (j < 0) j = fmt.length;
      push('lit', fmt.slice(i + 1, j));
      i = j + 1;
    } else if (c === '\\') {
      push('lit', fmt[i + 1] || '');
      i += 2;
    } else if (c === '[') {
      let j = fmt.indexOf(']', i);
      if (j < 0) j = fmt.length;
      i = j + 1; // 条件/颜色，忽略
    } else if (/^AM\/PM/i.test(fmt.slice(i, i + 5))) {
      push('ampm', fmt.substr(i, 5)); i += 5;
    } else if (/^A\/P/i.test(fmt.slice(i, i + 3))) {
      push('ampm', fmt.substr(i, 3)); i += 3;
    } else if (/[yYmMdDhHsS]/.test(c)) {
      const lc = c.toLowerCase();
      let j = i;
      while (j < fmt.length && fmt[j].toLowerCase() === lc) j++;
      push(lc, fmt.slice(i, j));
      i = j;
    } else if (/[0#?]/.test(c)) {
      let j = i;
      while (j < fmt.length && /[0#?.,]/.test(fmt[j])) j++;
      push('num', fmt.slice(i, j));
      i = j;
    } else {
      let j = i;
      while (j < fmt.length && !/[yYmMdDhHsS0#?"\\[]/.test(fmt[j]) &&
        !/^AM\/PM/i.test(fmt.slice(j, j + 5)) && !/^A\/P/i.test(fmt.slice(j, j + 3))) j++;
      if (j === i) j++;
      push('lit', fmt.slice(i, j));
      i = j;
    }
  }
  return toks;
}

function isDateFmt(toks) {
  let hasYmd = false, hasNum = false;
  for (const t of toks) {
    if (t.t === 'y' || t.t === 'd') hasYmd = true;
    if (t.t === 'm' || t.t === 'h' || t.t === 's') hasYmd = true;
    if (t.t === 'num') hasNum = true;
  }
  if (hasNum && !hasYmd) return false;
  return hasYmd;
}

export function formatNumber(value, fmt) {
  if (fmt == null || fmt === '' || fmt === 'General') {
    if (typeof value === 'number') {
      if (!isFinite(value)) return String(value);
      if (Number.isInteger(value)) return String(value);
      return String(parseFloat(value.toPrecision(12)));
    }
    return String(value);
  }
  if (typeof value !== 'number') return String(value);

  let f = String(fmt).split(';')[0];
  f = f.replace(/\[[^\]]*\]/g, '');
  if (f.trim() === '') return String(value);
  const toks = tokenize(f);
  if (isDateFmt(toks)) return formatDate(value, toks);
  return formatNumeric(value, f);
}

function formatNumeric(value, f) {
  if (f.includes('%')) {
    const m = f.match(/[0#?]*(?:\.([0#]+))?/);
    const dec = (f.match(/\.([0#]+)/) || [, ''])[1].length;
    return (value * 100).toFixed(dec) + '%';
  }
  if (/E[+-]/i.test(f)) return value.toExponential(2).replace(/e([+-])(\d)$/i, 'E$10$2').toUpperCase();
  let dec = 0;
  const dm = f.match(/\.([0#]+)/);
  if (dm) dec = dm[1].length;
  let out = Math.abs(value).toFixed(dec);
  if (f.includes(',')) {
    const [i, d] = out.split('.');
    out = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d ? '.' + d : '');
  }
  const sym = (f.match(/[$¥€£]/) || [''])[0];
  if (sym) out = sym + out;
  return (value < 0 ? '-' : '') + out;
}

function formatDate(serial, toks) {
  const d = excelSerialToDate(serial);
  const totalSec = Math.round((serial - Math.floor(serial)) * 86400);
  const hh = Math.floor(totalSec / 3600) % 24;
  const mi = Math.floor(totalSec / 60) % 60;
  const ss = totalSec % 60;
  let out = '';
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    switch (t.t) {
      case 'y': out += t.s.length <= 2 ? pad(d.getUTCFullYear() % 100, 2) : d.getUTCFullYear(); break;
      case 'm': {
        // 判断月还是分：看前一个 token 是 h，或后一个 token 是 s
        let isMin = false;
        for (let p = k - 1; p >= 0; p--) { if (toks[p].t === 'h') { isMin = true; break; } if (toks[p].t !== 'lit') break; }
        if (!isMin) for (let p = k + 1; p < toks.length; p++) { if (toks[p].t === 's') { isMin = true; break; } if (toks[p].t !== 'lit') break; }
        if (isMin) out += t.s.length >= 2 ? pad(mi, 2) : String(mi);
        else if (t.s.length >= 4) out += MONTHS[d.getUTCMonth()];
        else if (t.s.length === 3) out += MONTHS[d.getUTCMonth()].slice(0, 3);
        else out += t.s.length >= 2 ? pad(d.getUTCMonth() + 1, 2) : String(d.getUTCMonth() + 1);
        break;
      }
      case 'd':
        if (t.s.length >= 4) out += DAYS[d.getUTCDay()];
        else if (t.s.length === 3) out += DAYS[d.getUTCDay()].slice(0, 3);
        else out += t.s.length >= 2 ? pad(d.getUTCDate(), 2) : String(d.getUTCDate());
        break;
      case 'h': {
        let v = hh;
        // 若格式含 AM/PM 且是 12 小时制
        if (toks.some((x) => x.t === 'ampm')) { v = hh % 12; if (v === 0) v = 12; }
        out += t.s.length >= 2 ? pad(v, 2) : String(v);
        break;
      }
      case 's': out += t.s.length >= 2 ? pad(ss, 2) : String(ss); break;
      case 'ampm': out += hh >= 12 ? 'PM' : 'AM'; break;
      case 'num': {
        // 形如 .0 / .00 -> 秒的小数
        const m = t.s.match(/\.(0+)/);
        if (m) {
          const frac = (serial - Math.floor(serial)) * 86400;
          const sub = (frac - Math.floor(frac)).toFixed(m[1].length).slice(1);
          out += sub;
        }
        break;
      }
      default: out += t.s;
    }
  }
  return out;
}

/** 中文字体族映射（安卓仅有 Noto Sans/Serif CJK） */
export function fontFamily(name) {
  if (!name) return null;
  const n = String(name).toLowerCase();
  if (/mono|consol|courier|仿宋|fangsong|等线.*code/.test(n)) return 'monospace';
  if (/simsun|宋体|song|serif|times|georgia|cambria|book|楷|kai|仿宋|fz|stsong|ming/.test(n)) return 'serif';
  if (/simhei|黑体|yahei|微软雅黑|dengxian|等线|sans|arial|calibri|helvet|noto|segoe|verdana/.test(n)) return 'sans-serif';
  return null;
}

/** data: URL 生成（Node 测试用） */
export function bytesToDataUrl(bytes, mime) {
  let bin = '';
  const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  }
  return `data:${mime};base64,` + (typeof btoa === 'function' ? btoa(bin) : Buffer.from(bin, 'binary').toString('base64'));
}

export function mimeOf(name) {
  const e = name.split('.').pop().toLowerCase();
  return {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
    bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp',
    emf: 'image/emf', wmf: 'image/wmf', tif: 'image/tiff', tiff: 'image/tiff',
  }[e] || 'application/octet-stream';
}

export const IMG_EXT = /\.(png|jpe?g|gif|bmp|webp|svg)$/i;
