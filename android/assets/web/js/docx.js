/* Word (.docx) -> HTML 渲染器：向量文字，任意缩放不模糊 */
import * as X from './oxml.js';
import { OoxmlZip, parseRels, resolveRel } from './zipx.js';

const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff',
  blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080',
  darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', gray: '#808080', lightGray: '#c0c0c0',
  black: '#000000', white: '#ffffff',
};

const DEFAULTS = {
  pageW: 11906, pageH: 16838,   // A4, twips
  marT: 1440, marR: 1800, marB: 1440, marL: 1800,
};

function num2(v, d = 0) { const f = parseFloat(v); return isNaN(f) ? d : f; }

class StyleSheet {
  constructor(stylesXml) {
    this.byId = new Map();
    this.defaultPara = null;
    this.defaultRun = null;
    if (!stylesXml) return;
    const dd = X.el(stylesXml, 'docDefaults');
    if (dd) {
      const rPr = X.el(dd, 'rPr');
      const pPr = X.el(dd, 'pPr');
      if (rPr) this.defaultRun = rPr;
      if (pPr) this.defaultPara = pPr;
      const rFonts = X.el(rPr, 'rFonts');
      this.defaultFont = rFonts ? (X.attr(rFonts, 'eastAsia') || X.attr(rFonts, 'ascii')) : null;
      const sz = X.el(rPr, 'sz');
      this.defaultSize = sz ? num2(X.attr(sz, 'val'), 21) : 21;
    }
    for (const s of X.els(stylesXml, 'style')) {
      const id = X.attr(s, 'styleId');
      if (!id) continue;
      this.byId.set(id, {
        id,
        type: X.attr(s, 'type'),
        basedOn: X.attr(X.child(s, 'basedOn'), 'val'),
        pPr: X.child(s, 'pPr'),
        rPr: X.child(s, 'rPr'),
        tblPr: X.child(s, 'tblPr'),
      });
    }
  }
  chain(id) {
    const out = [];
    const seen = new Set();
    let cur = id;
    while (cur && this.byId.has(cur) && !seen.has(cur)) {
      seen.add(cur);
      const s = this.byId.get(cur);
      out.push(s);
      cur = s.basedOn;
    }
    return out.reverse(); // 从根到叶
  }
}

/** 读取 run 属性 */
function readRunProps(node, base) {
  const p = base ? { ...base } : {};
  if (!node) return p;
  const g = (n) => X.child(node, n);
  const b = g('b'); if (b) p.b = X.attr(b, 'val') !== '0' && X.attr(b, 'val') !== 'false';
  const i = g('i'); if (i) p.i = X.attr(i, 'val') !== '0' && X.attr(i, 'val') !== 'false';
  const u = g('u'); if (u) p.u = X.attr(u, 'val') !== 'none';
  const st = g('strike'); if (st) p.strike = true;
  const sz = g('sz'); if (sz) p.sz = num2(X.attr(sz, 'val'), 21);
  const c = g('color'); if (c) { const v = X.color(X.attr(c, 'val')); if (v) p.color = v; else if (X.attr(c, 'val') === 'auto') p.color = null; }
  const hl = g('highlight'); if (hl) { const v = HIGHLIGHT[X.attr(hl, 'val')]; if (v) p.bg = v; }
  const shd = g('shd'); if (shd) { const v = X.color(X.attr(shd, 'fill')); if (v) p.bg = v; }
  const rf = g('rFonts');
  if (rf) {
    const f = X.attr(rf, 'eastAsia') || X.attr(rf, 'ascii') || X.attr(rf, 'hAnsi');
    if (f) p.font = f;
  }
  const va = g('vertAlign');
  if (va) { const v = X.attr(va, 'val'); p.va = (v === 'superscript' || v === 'subscript') ? v : null; }
  const sp = g('spacing'); if (sp) p.ls = num2(X.attr(sp, 'val'), 0);
  const caps = g('caps'); if (caps && X.attr(caps, 'val') !== '0') p.caps = true;
  return p;
}

function runCss(p, defSize) {
  const st = [];
  if (p.font) { const f = X.fontFamily(p.font); if (f) st.push('font-family:' + f); }
  const sz = p.sz != null ? p.sz : defSize;
  if (sz) st.push('font-size:' + (sz * 2 / 3).toFixed(2) + 'px');
  if (p.b) st.push('font-weight:700');
  if (p.i) st.push('font-style:italic');
  if (p.caps) st.push('text-transform:uppercase');
  const deco = [];
  if (p.u) deco.push('underline');
  if (p.strike) deco.push('line-through');
  if (deco.length) st.push('text-decoration:' + deco.join(' '));
  if (p.color) st.push('color:' + p.color);
  if (p.bg) st.push('background:' + p.bg);
  if (p.va) st.push('vertical-align:' + (p.va === 'superscript' ? 'super' : 'sub'), 'font-size:0.7em');
  if (p.ls) st.push('letter-spacing:' + (p.ls / 15).toFixed(2) + 'px');
  return st.join(';');
}

function readParaProps(node, base) {
  const p = base ? { ...base } : {};
  if (!node) return p;
  const g = (n) => X.child(node, n);
  const jc = g('jc'); if (jc) p.align = { both: 'justify', distribute: 'justify', start: 'left', end: 'right' }[X.attr(jc, 'val')] || X.attr(jc, 'val');
  const ind = g('ind');
  if (ind) {
    if (X.attr(ind, 'left') != null) p.left = num2(X.attr(ind, 'left'));
    else if (X.attr(ind, 'start') != null) p.left = num2(X.attr(ind, 'start'));
    if (X.attr(ind, 'right') != null) p.right = num2(X.attr(ind, 'right'));
    if (X.attr(ind, 'firstLine') != null) { p.firstLine = num2(X.attr(ind, 'firstLine')); p.hanging = 0; }
    if (X.attr(ind, 'hanging') != null) { p.hanging = num2(X.attr(ind, 'hanging')); p.firstLine = 0; }
  }
  const sp = g('spacing');
  if (sp) {
    if (X.attr(sp, 'before') != null) p.before = num2(X.attr(sp, 'before'));
    if (X.attr(sp, 'after') != null) p.after = num2(X.attr(sp, 'after'));
    if (X.attr(sp, 'line') != null) { p.line = num2(X.attr(sp, 'line')); p.lineRule = X.attr(sp, 'lineRule') || 'auto'; }
  }
  const shd = g('shd'); if (shd) { const v = X.color(X.attr(shd, 'fill')); if (v) p.bg = v; }
  if (g('pageBreakBefore')) p.pageBreakBefore = true;
  const bdr = g('pBdr');
  if (bdr) p.bdr = bdr;
  const numPr = g('numPr');
  if (numPr) p.numPr = numPr;
  return p;
}

function paraCss(p, inTable) {
  const st = [];
  if (p.align) st.push('text-align:' + p.align);
  const left = (p.left || 0) / 15;
  const right = (p.right || 0) / 15;
  if (left) st.push('margin-left:' + left.toFixed(2) + 'px');
  if (right) st.push('margin-right:' + right.toFixed(2) + 'px');
  if (p.firstLine) st.push('text-indent:' + (p.firstLine / 15).toFixed(2) + 'px');
  if (p.hanging) st.push('text-indent:' + (-p.hanging / 15).toFixed(2) + 'px');
  const before = (p.before || 0) / 15, after = (p.after || 0) / 15;
  if (before) st.push('margin-top:' + before.toFixed(2) + 'px');
  if (after) st.push('margin-bottom:' + after.toFixed(2) + 'px');
  if (p.line) {
    if (p.lineRule === 'auto') st.push('line-height:' + (p.line / 240).toFixed(3));
    else st.push('line-height:' + (p.line / 15).toFixed(2) + 'px');
  }
  if (p.bg) st.push('background:' + p.bg);
  if (p.bdr) {
    const sides = { top: 'top', bottom: 'bottom', left: 'left', right: 'right' };
    for (const k of Object.keys(sides)) {
      const b = X.child(p.bdr, k);
      if (!b) continue;
      const v = X.attr(b, 'val');
      if (!v || v === 'none' || v === 'nil') continue;
      const c = X.color(X.attr(b, 'color')) || '#000';
      const w = Math.max(1, Math.round(num2(X.attr(b, 'sz'), 4) / 8 * 96 / 72));
      st.push('border-' + sides[k] + ':' + w + 'px solid ' + c);
      const sp = num2(X.attr(b, 'space'), 0) / 15;
      if (sp) st.push('padding-' + sides[k] + ':' + sp.toFixed(1) + 'px');
    }
  }
  return st.join(';');
}

const ROMAN = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
  [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
function toRoman(n) { let s = ''; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s.toLowerCase(); }
function toLetters(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(97 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
const CN_NUM = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
function toChinese(n) {
  if (n <= 10) return n === 10 ? '十' : CN_NUM[n];
  if (n < 20) return '十' + CN_NUM[n - 10];
  if (n < 100) return CN_NUM[Math.floor(n / 10)] + '十' + (n % 10 ? CN_NUM[n % 10] : '');
  return String(n);
}

export function renderDocx(bytes, opts = {}) {
  const zip = new OoxmlZip(bytes);
  const docXml = zip.text('word/document.xml');
  if (!docXml) throw new Error('不是有效的 .docx 文件');
  const doc = X.parseXml(docXml);
  const stylesXmlSrc = zip.text('word/styles.xml');
  const stylesXml = stylesXmlSrc ? X.parseXml(stylesXmlSrc) : null;
  const ss = new StyleSheet(stylesXml);
  const numXmlSrc = zip.text('word/numbering.xml');
  const numbering = numXmlSrc ? X.parseXml(numXmlSrc) : null;
  const rels = parseRels(zip.text('word/_rels/document.xml.rels') ? X.parseXml(zip.text('word/_rels/document.xml.rels')) : null);

  // 页面尺寸
  const body = X.el(doc, 'body');
  const bodySect = body ? X.child(body, 'sectPr') : null;
  const sect = bodySect || X.el(doc, 'sectPr');
  let pageW = DEFAULTS.pageW, pageH = DEFAULTS.pageH, marT = DEFAULTS.marT, marR = DEFAULTS.marR, marB = DEFAULTS.marB, marL = DEFAULTS.marL;
  if (sect) {
    const pgSz = X.child(sect, 'pgSz');
    if (pgSz) { pageW = num2(X.attr(pgSz, 'w'), pageW); pageH = num2(X.attr(pgSz, 'h'), pageH); }
    const pgMar = X.child(sect, 'pgMar');
    if (pgMar) {
      marT = num2(X.attr(pgMar, 'top'), marT);
      marR = num2(X.attr(pgMar, 'right'), marR);
      marB = num2(X.attr(pgMar, 'bottom'), marB);
      marL = num2(X.attr(pgMar, 'left'), marL);
    }
  }

  // 编号定义
  const abstractNums = new Map();
  const nums = new Map();
  if (numbering) {
    for (const a of X.els(numbering, 'abstractNum')) {
      const id = X.attr(a, 'abstractNumId');
      const lvls = [];
      for (const l of X.children(a, 'lvl')) {
        lvls[num2(X.attr(l, 'ilvl'))] = {
          numFmt: X.attr(X.child(l, 'numFmt'), 'val') || 'decimal',
          start: num2(X.attr(X.child(l, 'start'), 'val'), 1),
          text: X.attr(X.child(l, 'lvlText'), 'val') || '',
          pPr: X.child(l, 'pPr'),
          rPr: X.child(l, 'rPr'),
        };
      }
      abstractNums.set(id, lvls);
    }
    for (const n of X.els(numbering, 'num')) {
      const numId = X.attr(n, 'numId');
      const a = X.attr(X.child(n, 'abstractNumId'), 'val');
      const overrides = new Map();
      for (const o of X.children(n, 'lvlOverride')) {
        overrides.set(X.attr(o, 'ilvl'), num2(X.attr(X.child(o, 'startOverride'), 'val'), null));
      }
      nums.set(numId, { abstractId: a, overrides });
    }
  }
  const counters = new Map(); // abstractId -> [counts]

  function markerFor(numId, ilvl) {
    const n = nums.get(numId);
    if (!n) return null;
    const lvls = abstractNums.get(n.abstractId);
    if (!lvls) return null;
    const lvl = lvls[ilvl];
    if (!lvl) return null;
    let c = counters.get(n.abstractId);
    if (!c) { c = []; counters.set(n.abstractId, c); }
    c[ilvl] = (c[ilvl] == null ? lvl.start - 1 : c[ilvl]) + 1;
    for (let k = ilvl + 1; k < 9; k++) c[k] = null;

    const fmt = lvl.numFmt;
    if (fmt === 'bullet' || fmt === 'none') {
      return { text: lvl.text && lvl.text !== '' ? lvl.text : '•', bullet: true, lvl };
    }
    let text = lvl.text || '%1.';
    text = text.replace(/%(\d)/g, (m, d) => {
      const lv = parseInt(d, 10) - 1;
      const v = c[lv] != null ? c[lv] : (lvls[lv] ? lvls[lv].start : 1);
      const f2 = lvls[lv] ? lvls[lv].numFmt : 'decimal';
      return fmtNum(v, f2);
    });
    return { text, bullet: false, lvl };
  }

  function fmtNum(v, fmt) {
    switch (fmt) {
      case 'lowerLetter': return toLetters(v);
      case 'upperLetter': return toLetters(v).toUpperCase();
      case 'lowerRoman': return toRoman(v);
      case 'upperRoman': return toRoman(v).toUpperCase();
      case 'chineseCounting': case 'chineseCountingThousand': case 'chineseLegalSimplified': case 'ideographDigital': return toChinese(v);
      case 'decimalZero': return v < 10 ? '0' + v : String(v);
      case 'decimalEnclosedCircle': return v >= 1 && v <= 20 ? String.fromCharCode(0x2460 + v - 1) : String(v);
      case 'decimalFullWidth': return String(v).replace(/\d/g, (d) => String.fromCharCode(0xFF10 + +d));
      default: return String(v);
    }
  }

  function mediaImg(rid, fallback) {
    const r = rels[rid];
    if (!r) return null;
    if (r.external) return r.target;
    const path = resolveRel('word', r);
    return zip.url(path);
  }

  function renderRun(r, baseRun, ctx) {
    const p = readRunProps(X.child(r, 'rPr'), baseRun);
    const out = [];
    for (const c of X.kids(r)) {
      switch (c.localName) {
        case 't': {
          let t = c.textContent || '';
          t = t.replace(/\u00a0/g, ' ').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
          if (t) out.push(X.esc(t));
          break;
        }
        case 'tab': out.push('<span style="display:inline-block;width:2em"></span>'); break;
        case 'br':
          if (X.attr(c, 'type') === 'page') out.push('<span class="dw-pgb"></span>');
          else out.push('<br>');
          break;
        case 'cr': out.push('<br>'); break;
        case 'noBreakHyphen': out.push('-'); break;
        case 'sym': break;
        case 'drawing': out.push(renderDrawing(c, ctx)); break;
        case 'pict': out.push(renderPict(c, ctx)); break;
        case 'object': out.push(renderPict(c, ctx)); break;
        default: break;
      }
    }
    const html = out.join('');
    if (!html) return '';
    const css = runCss(p, ss.defaultSize);
    return css ? `<span style="${css}">${html}</span>` : html;
  }

  function renderDrawing(d, ctx) {
    const blip = X.el(d, 'blip');
    const ext = X.el(d, 'extent') || X.el(d, 'ext');
    let cx = ext ? num2(X.attr(ext, 'cx'), 0) : 0;
    let cy = ext ? num2(X.attr(ext, 'cy'), 0) : 0;
    if (blip) {
      const rid = X.attr(blip, 'embed') || X.attr(blip, 'link');
      const src = mediaImg(rid);
      if (src) {
        const w = cx ? cx / X.EMU : null;
        const h = cy ? cy / X.EMU : null;
        const style = `${w ? 'width:' + w.toFixed(1) + 'px;' : ''}${h && !w ? 'height:' + h.toFixed(1) + 'px;' : ''}max-width:100%;object-fit:contain;vertical-align:middle`;
        return `<img src="${src}" style="${style}" alt="">`;
      }
    }
    const txbx = X.el(d, 'txbxContent');
    if (txbx) {
      const w = cx ? cx / X.EMU : 200;
      return `<span style="display:inline-block;width:${w.toFixed(1)}px;vertical-align:top">${renderBlocks(txbx, ctx)}</span>`;
    }
    return '';
  }

  function renderPict(p, ctx) {
    const imd = X.el(p, 'imagedata');
    if (imd) {
      const rid = X.attr(imd, 'id') || X.attr(imd, 'embed');
      const src = mediaImg(rid);
      if (src) {
        const shape = X.el(p, 'shape');
        let w = 0, h = 0;
        const style = shape ? X.attr(shape, 'style') || '' : '';
        const mw = style.match(/width:\s*([\d.]+)pt/); if (mw) w = parseFloat(mw[1]) * X.PT;
        const mh = style.match(/height:\s*([\d.]+)pt/); if (mh) h = parseFloat(mh[1]) * X.PT;
        const s = `${w ? 'width:' + w.toFixed(1) + 'px;' : ''}${h ? 'height:' + h.toFixed(1) + 'px;' : ''}max-width:100%;object-fit:contain;vertical-align:middle`;
        return `<img src="${src}" style="${s}" alt="">`;
      }
    }
    return '';
  }

  function renderParagraph(p, baseRunIn, ctx) {
    const pPrNode = X.child(p, 'pPr');
    const pStyleId = X.attr(X.child(pPrNode, 'pStyle'), 'val');
    let paraBase = {};
    let runBase = readRunProps(ss.defaultRun, {});
    if (pStyleId) {
      for (const s of ss.chain(pStyleId)) {
        paraBase = readParaProps(s.pPr, paraBase);
        runBase = readRunProps(s.rPr, runBase);
      }
    }
    const para = readParaProps(pPrNode, paraBase);
    const runBaseFinal = baseRunIn ? { ...runBase, ...baseRunIn } : runBase;

    // 编号
    let marker = null;
    const numPr = para.numPr;
    if (numPr) {
      const ilvl = num2(X.attr(X.child(numPr, 'ilvl'), 'val'), 0);
      const numId = X.attr(X.child(numPr, 'numId'), 'val');
      if (numId && numId !== '0') marker = markerFor(numId, ilvl);
    }

    const parts = [];
    for (const c of X.kids(p)) {
      switch (c.localName) {
        case 'r': parts.push(renderRun(c, runBaseFinal, ctx)); break;
        case 'hyperlink':
          for (const r of X.children(c, 'r')) parts.push(renderRun(r, { ...runBaseFinal, color: '#1a5fb4', u: true }, ctx));
          break;
        case 'ins': for (const r of X.children(c, 'r')) parts.push(renderRun(r, runBaseFinal, ctx)); break;
        case 'smartTag': case 'sdt': case 'sdtContent':
          for (const r of X.els(c, 'r')) parts.push(renderRun(r, runBaseFinal, ctx));
          break;
        default: break;
      }
    }
    let inner = parts.join('');
    if (/class="dw-pgb"/.test(inner)) {
      inner = inner.replace(/<span class="dw-pgb"><\/span>/g, '');
      ctx.pendingBreak = true;
    }
    if (!inner.trim()) inner += '<br>';
    if (ctx.pendingBreak) { ctx.out.push('<div class="dw-pgb"></div>'); ctx.pendingBreak = false; }

    let css = paraCss(para, ctx.inTable);
    if (marker) {
      const lvlInd = marker.lvl && marker.lvl.pPr ? X.child(marker.lvl.pPr, 'ind') : null;
      const lInd = lvlInd ? num2(X.attr(lvlInd, 'left'), 420) : 420;
      const hInd = lvlInd ? num2(X.attr(lvlInd, 'hanging'), 420) : 420;
      const leftPx = (para.left != null ? para.left : lInd) / 15;
      const hangPx = hInd / 15;
      const mStyle = marker.bullet
        ? `display:inline-block;width:${hangPx.toFixed(1)}px;text-indent:0;`
        : `display:inline-block;width:${hangPx.toFixed(1)}px;padding-right:4px;box-sizing:border-box;text-indent:0;`;
      css += `;padding-left:${leftPx.toFixed(1)}px;text-indent:${(-hangPx).toFixed(1)}px`;
      inner = `<span style="${mStyle}">${marker.bullet ? '•' : X.esc(marker.text)}</span>` + inner;
    }
    const tag = marker ? 'div' : 'p';
    ctx.out.push(`<${tag} class="dw-p" style="${css}">${inner}</${tag}>`);
  }

  function borderDef(node) {
    if (!node) return null;
    const v = X.attr(node, 'val');
    if (!v || v === 'none' || v === 'nil') return { none: true };
    const c = X.color(X.attr(node, 'color')) || '#333333';
    const w = Math.max(1, Math.round(num2(X.attr(node, 'sz'), 4) / 8 * 96 / 72));
    return { w, c };
  }

  function tableBorders(tblPr) {
    const out = {};
    const styleId = X.attr(X.child(tblPr, 'tblStyle'), 'val');
    const chains = styleId ? ss.chain(styleId) : [];
    for (const st of chains) {
      const tp = st.tblPr;
      const bd = tp ? X.child(tp, 'tblBorders') : null;
      if (!bd) continue;
      for (const k of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
        const b = X.child(bd, k);
        if (b) out[k] = borderDef(b);
      }
    }
    const direct = X.child(tblPr, 'tblBorders');
    if (direct) {
      for (const k of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
        const b = X.child(direct, k);
        if (b) out[k] = borderDef(b);
      }
    }
    return out;
  }

  function bdCss(side, def) {
    if (!def) return `border-${side}:none`;
    if (def.none) return `border-${side}:none`;
    return `border-${side}:${def.w}px solid ${def.c}`;
  }

  function renderTable(tbl, ctx) {
    const tblPr = X.child(tbl, 'tblPr');
    const st = ['border-collapse:collapse', 'table-layout:fixed', 'box-sizing:border-box'];
    let width = null, align = 'left';
    const tblW = tblPr ? X.child(tblPr, 'tblW') : null;
    if (tblW) {
      const t = X.attr(tblW, 'type');
      const w = num2(X.attr(tblW, 'w'));
      if (t === 'dxa' && w) width = w / 15;
      else if (t === 'pct') { const pcent = w / 50; st.push('width:' + pcent.toFixed(2) + '%'); }
    }
    const jc = tblPr ? X.child(tblPr, 'jc') : null;
    if (jc) { const v = X.attr(jc, 'val'); if (v === 'center') align = 'center'; else if (v === 'end' || v === 'right') align = 'right'; }
    if (align === 'center') st.push('margin-left:auto', 'margin-right:auto');
    else if (align === 'right') st.push('margin-left:auto');
    const tblInd = tblPr ? X.child(tblPr, 'tblInd') : null;
    if (tblInd && num2(X.attr(tblInd, 'w'))) st.push('margin-left:' + (num2(X.attr(tblInd, 'w')) / 15).toFixed(1) + 'px');

    const borders = tableBorders(tblPr);
    const hasBorder = Object.keys(borders).some((k) => borders[k] && !borders[k].none);

    const grid = X.child(tbl, 'tblGrid');
    const cols = [];
    if (grid) for (const gc of X.children(grid, 'gridCol')) cols.push(num2(X.attr(gc, 'w'), 100) / 15);

    let defMarL = 108 / 15, defMarR = 108 / 15, defMarT = 0, defMarB = 0;
    const cm = tblPr ? X.child(tblPr, 'tblCellMar') : null;
    if (cm) {
      const l = X.child(cm, 'left'), r = X.child(cm, 'right'), tp = X.child(cm, 'top'), bt = X.child(cm, 'bottom');
      if (l) defMarL = num2(X.attr(l, 'w'), 108) / 15;
      if (r) defMarR = num2(X.attr(r, 'w'), 108) / 15;
      if (tp) defMarT = num2(X.attr(tp, 'w'), 0) / 15;
      if (bt) defMarB = num2(X.attr(bt, 'w'), 0) / 15;
    }

    if (width) st.push('width:' + width.toFixed(1) + 'px');
    else if (cols.length) st.push('width:' + cols.reduce((a, b) => a + b, 0).toFixed(1) + 'px');
    else st.push('width:100%');

    const rows = X.children(tbl, 'tr');
    const rowsCells = rows.map((tr) => X.children(tr, 'tc'));
    let html = `<table class="dw-tbl" style="${st.join(';')}">`;
    if (cols.length) html += '<colgroup>' + cols.map((w) => `<col style="width:${w.toFixed(2)}px">`).join('') + '</colgroup>';
    html += '<tbody>';
    const occupied = [];

    rows.forEach((tr, ri) => {
      const trPr = X.child(tr, 'trPr');
      let rowCss = '';
      const trH = trPr ? X.child(trPr, 'trHeight') : null;
      if (trH) {
        const h = num2(X.attr(trH, 'val')) / 15;
        if (h) rowCss = 'height:' + h.toFixed(1) + 'px';
      }
      const isFirstRow = ri === 0, isLastRow = ri === rows.length - 1;
      html += `<tr style="${rowCss}">`;
      let col = 0;
      const cells = X.children(tr, 'tc');
      for (const tc of cells) {
        while (occupied[ri] && occupied[ri][col]) col++;
        const tcPr = X.child(tc, 'tcPr');
        const span = tcPr ? num2(X.attr(X.child(tcPr, 'gridSpan'), 'val'), 1) : 1;
        const vMergeNode = tcPr ? X.child(tcPr, 'vMerge') : null;
        let vMerge = null;
        if (vMergeNode) vMerge = X.attr(vMergeNode, 'val') === 'restart' ? 'restart' : 'continue';

        if (vMerge === 'continue') {
          for (let k = 0; k < span; k++) { if (!occupied[ri]) occupied[ri] = []; occupied[ri][col + k] = true; }
          col += span;
          continue;
        }
        let rowspan = 1;
        if (vMerge === 'restart') {
          let r2 = ri + 1;
          while (r2 < rows.length) {
            const tcs = rowsCells[r2];
            let c2 = 0, found = null;
            for (const t2 of tcs) {
              while (occupied[r2] && occupied[r2][c2]) c2++;
              if (c2 === col) { found = t2; break; }
              const p2 = X.child(t2, 'tcPr'); const s2 = p2 ? num2(X.attr(X.child(p2, 'gridSpan'), 'val'), 1) : 1;
              c2 += s2;
            }
            if (!found) break;
            const p2 = X.child(found, 'tcPr');
            const vm = p2 ? X.child(p2, 'vMerge') : null;
            if (vm && X.attr(vm, 'val') !== 'restart') {
              rowspan++;
              if (!occupied[r2]) occupied[r2] = [];
              for (let k = 0; k < span; k++) occupied[r2][col + k] = true;
            } else break;
            r2++;
          }
        }

        const tdCss = [];
        const tcMar = tcPr ? X.child(tcPr, 'tcMar') : null;
        let mL = defMarL, mR = defMarR, mT = defMarT, mB = defMarB;
        if (tcMar) {
          for (const k of ['left', 'right', 'top', 'bottom']) {
            const n = X.child(tcMar, k);
            if (!n) continue;
            const v = num2(X.attr(n, 'w'), 0) / 15;
            if (k === 'left') mL = v; else if (k === 'right') mR = v;
            else if (k === 'top') mT = v; else mB = v;
          }
        }
        tdCss.push(`padding:${mT.toFixed(1)}px ${mR.toFixed(1)}px ${mB.toFixed(1)}px ${mL.toFixed(1)}px`);

        // 边框：单元格自定义 > 表格样式 > 表格直接定义
        const tcB = tcPr ? X.child(tcPr, 'tcBorders') : null;
        const eff = {
          top: (tcB && X.child(tcB, 'top') ? borderDef(X.child(tcB, 'top')) : (isFirstRow ? borders.top : borders.insideH)),
          bottom: (tcB && X.child(tcB, 'bottom') ? borderDef(X.child(tcB, 'bottom')) : (isLastRow ? borders.bottom : borders.insideH)),
          left: (tcB && X.child(tcB, 'left') ? borderDef(X.child(tcB, 'left')) : (col === 0 ? borders.left : borders.insideV)),
          right: (tcB && X.child(tcB, 'right') ? borderDef(X.child(tcB, 'right')) : (col + span >= 199 ? borders.right : borders.insideV)),
        };
        if (hasBorder || true) {
          for (const k of ['top', 'bottom', 'left', 'right']) tdCss.push(bdCss(k, eff[k]));
        }

        const shd = tcPr ? X.child(tcPr, 'shd') : null;
        if (shd) { const c = X.color(X.attr(shd, 'fill')); if (c) tdCss.push('background:' + c); }
        const va = tcPr ? X.child(tcPr, 'vAlign') : null;
        if (va) tdCss.push('vertical-align:' + ({ center: 'middle', bottom: 'bottom' }[X.attr(va, 'val')] || 'top'));

        const inner = [];
        const subCtx = { ...ctx, out: inner, inTable: true, pendingBreak: false };
        for (const el2 of X.kids(tc)) {
          if (el2.localName === 'p') renderParagraph(el2, null, subCtx);
          else if (el2.localName === 'tbl') inner.push(renderTable(el2, subCtx));
        }
        html += `<td${span > 1 ? ` colspan="${span}"` : ''}${rowspan > 1 ? ` rowspan="${rowspan}"` : ''} style="${tdCss.join(';')}">${inner.join('')}</td>`;
        if (!occupied[ri]) occupied[ri] = [];
        for (let k = 0; k < span; k++) occupied[ri][col + k] = true;
        col += span;
      }
      html += '</tr>';
    });
    html += '</tbody></table>';
    return html;
  }

  function renderBlocks(node, ctx) {
    const out = [];
    const c2 = { ...ctx, out, inTable: false };
    for (const el2 of X.kids(node)) {
      if (el2.localName === 'p') renderParagraph(el2, null, c2);
      else if (el2.localName === 'tbl') c2.out.push(renderTable(el2, c2));
      else if (el2.localName === 'sdt') {
        const content = X.child(el2, 'sdtContent');
        if (content) c2.out.push(renderBlocks(content, c2));
      } else if (el2.localName === 'sectPr') { /* 分节符 */ }
    }
    return out.join('');
  }

  const ctx = { out: [], inTable: false, pendingBreak: false };
  let html = renderBlocks(body || doc.documentElement, ctx);

  const pageWpx = pageW / 15, pageHpx = pageH / 15;
  const padT = marT / 15, padR = marR / 15, padB = marB / 15, padL = marL / 15;
  const pageCount = (html.match(/class="dw-pgb"/g) || []).length + 1;

  const sheet = `<div class="doc-sheet" style="width:${pageWpx.toFixed(1)}px;padding:${padT.toFixed(1)}px ${padR.toFixed(1)}px ${padB.toFixed(1)}px ${padL.toFixed(1)}px">${html}</div>`;

  return {
    kind: 'docx',
    html: sheet,
    pageW: pageWpx, pageH: pageHpx,
    padT, padR, padB, padL,
    pageCount,
    flow: true,
    cleanup: () => zip.revoke(),
    searchText: textOfDocx(doc),
  };
}

function textOfDocx(doc) {
  // 抽全文（用于搜索/预览）
  const paras = X.els(doc, 'p');
  const out = [];
  for (const p of paras) {
    const t = Array.from(p.getElementsByTagNameNS('*', 't')).map((x) => x.textContent).join('');
    if (t.trim()) out.push(t.trim());
    if (out.length > 400) break;
  }
  return out.join('\n');
}
