/* Excel (.xlsx) -> HTML 渲染器 */
import * as X from './oxml.js';
import { OoxmlZip, parseRels, resolveRel } from './zipx.js';

const BUILTIN_DATE_FMTS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34,
  35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

/* Excel 内置数字格式（id < 164） */
const BUILTIN_FMTS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '¥#,##0;[Red]¥-#,##0', 6: '¥#,##0;[Red]¥-#,##0', 7: '¥#,##0.00;[Red]¥-#,##0.00',
  8: '¥#,##0.00;[Red]¥-#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
  12: '# ?/?', 13: '# ??/??',
  14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};

function colToNum(ref) {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c >= 65 && c <= 90) n = n * 26 + (c - 64);
    else break;
  }
  return n;
}
function numToCol(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function splitRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  return { col: colToNum(m[1]), row: parseInt(m[2], 10) };
}

class Styles {
  constructor(xml) {
    this.numFmts = new Map();
    this.fonts = [];
    this.fills = [];
    this.borders = [];
    this.xfs = [];
    this.cellStyleXfs = [];
    if (!xml) return;
    const nf = X.el(xml, 'numFmts');
    if (nf) for (const n of X.children(nf, 'numFmt')) this.numFmts.set(parseInt(n.getAttribute('numFmtId')), n.getAttribute('formatCode'));

    const fontsNode = X.el(xml, 'fonts');
    if (fontsNode) {
      for (const f of X.children(fontsNode, 'font')) {
        const sz = X.el(f, 'sz');
        const b = X.el(f, 'b');
        const i = X.el(f, 'i');
        const u = X.el(f, 'u');
        const color = X.el(f, 'color');
        const name = X.el(f, 'name') || X.el(f, 'rFont');
        this.fonts.push({
          sz: sz ? parseFloat(sz.getAttribute('val')) : 11,
          b: !!b && X.attr(b, 'val') !== '0',
          i: !!i && X.attr(i, 'val') !== '0',
          u: !!u && X.attr(u, 'val') !== 'none',
          color: color ? (X.color(X.attr(color, 'rgb')) || X.themeColor(X.el(color, 'schemeClr')) ||
            (X.attr(color, 'theme') != null ? null : null)) : null,
          name: name ? name.getAttribute('val') : null,
          strike: !!X.el(f, 'strike'),
        });
      }
    }
    const fillsNode = X.el(xml, 'fills');
    if (fillsNode) {
      for (const f of X.children(fillsNode, 'fill')) {
        const pf = X.el(f, 'patternFill');
        let fg = null, type = null;
        if (pf) {
          type = pf.getAttribute('patternType');
          const fgc = X.el(pf, 'fgColor');
          if (fgc) fg = X.color(fgc.getAttribute('rgb')) || X.themeColor(X.el(fgc, 'schemeClr'));
          if (!fg && fgc && fgc.getAttribute('theme')) {
            const idx = parseInt(fgc.getAttribute('theme'), 10);
            fg = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'][idx] || null;
          }
        }
        this.fills.push({ fg, type });
      }
    }
    const bordersNode = X.el(xml, 'borders');
    if (bordersNode) {
      for (const b of X.children(bordersNode, 'border')) {
        const sides = {};
        for (const k of ['left', 'right', 'top', 'bottom']) {
          const s = X.el(b, k);
          if (!s) continue;
          const style = s.getAttribute('style');
          const c = X.el(s, 'color');
          sides[k] = style ? { style, color: c ? (X.color(c.getAttribute('rgb')) || X.themeColor(X.el(c, 'schemeClr'))) : null } : null;
        }
        this.borders.push(sides);
      }
    }
    const xfNode = X.el(xml, 'cellXfs');
    if (xfNode) {
      for (const xf of X.children(xfNode, 'xf')) {
        const al = X.el(xf, 'alignment');
        this.xfs.push({
          fontId: parseInt(xf.getAttribute('fontId') || '0', 10),
          fillId: parseInt(xf.getAttribute('fillId') || '0', 10),
          borderId: parseInt(xf.getAttribute('borderId') || '0', 10),
          numFmtId: parseInt(xf.getAttribute('numFmtId') || '0', 10),
          align: al ? {
            h: al.getAttribute('horizontal') || null,
            v: al.getAttribute('vertical') || null,
            wrap: al.getAttribute('wrapText') === '1',
            indent: parseInt(al.getAttribute('indent') || '0', 10),
          } : null,
        });
      }
    }
  }
  fmtCode(id) {
    if (this.numFmts.has(id)) return this.numFmts.get(id);
    if (BUILTIN_FMTS[id] != null) return BUILTIN_FMTS[id];
    return null;
  }
  isDate(id) {
    if (BUILTIN_DATE_FMTS.has(id)) return true;
    if (id < 164) return false;
    const c = this.fmtCode(id);
    if (!c) return false;
    const stripped = c.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
    return /[ymdhs]/i.test(stripped) && !/^[#0.,%\s]*$/.test(stripped);
  }
}

export function renderXlsx(bytes, opts = {}) {
  const zip = new OoxmlZip(bytes);
  const wbSrc = zip.text('xl/workbook.xml');
  if (!wbSrc) throw new Error('不是有效的 .xlsx 文件');
  const wb = X.parseXml(wbSrc);
  const rels = parseRels(zip.text('xl/_rels/workbook.xml.rels') ? X.parseXml(zip.text('xl/_rels/workbook.xml.rels')) : null);

  // 共享字符串
  const sst = [];
  const sstSrc = zip.text('xl/sharedStrings.xml');
  if (sstSrc) {
    const sstXml = X.parseXml(sstSrc);
    for (const si of X.children(X.el(sstXml, 'sst'), 'si')) {
      let t = '';
      const direct = X.child(si, 't');
      if (direct) t = direct.textContent;
      else for (const r of X.children(si, 'r')) { const tt = X.child(r, 't'); if (tt) t += tt.textContent; }
      if (X.child(si, 'rPh')) { /* 注音忽略 */ }
      sst.push(t);
    }
  }
  const styles = new Styles(zip.text('xl/styles.xml') ? X.parseXml(zip.text('xl/styles.xml')) : null);

  const sheetNodes = X.els(wb, 'sheet');
  const sheets = [];
  for (const s of sheetNodes) {
    const rid = s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ||
      s.getAttribute('r:id');
    const rel = rels[rid];
    if (!rel) continue;
    const path = resolveRel('xl', rel);
    sheets.push({ name: s.getAttribute('name'), path, state: s.getAttribute('state') || 'visible' });
  }

  function renderSheet(sh) {
    const src = zip.text(sh.path);
    if (!src) return { html: '<div class="xls-empty">工作表为空</div>', w: 400, h: 200, rows: 0 };
    const xml = X.parseXml(src);
    const dataNode = X.el(xml, 'sheetData');
    const rowsXml = X.children(dataNode, 'row');

    // 列宽
    const colW = new Map();
    let defaultW = 8.43;
    const colsNode = X.el(xml, 'cols');
    if (colsNode) {
      for (const c of X.children(colsNode, 'col')) {
        const mn = parseInt(c.getAttribute('min')), mx = parseInt(c.getAttribute('max'));
        const w = c.getAttribute('width');
        const cw = c.getAttribute('customWidth') === '1';
        if (c.getAttribute('width') != null) {
          for (let i = mn; i <= Math.min(mx, mn + 2000); i++) colW.set(i, cw ? parseFloat(w) : parseFloat(w));
        }
      }
    }
    const widthPx = (w) => Math.round(Math.min(w, 255) * 7 + 5);

    // 尺寸
    const dim = X.el(xml, 'dimension');
    let maxCol = 0, maxRow = 0;
    const cellSeq = [];
    for (const r of rowsXml) {
      const rIdx = parseInt(r.getAttribute('r') || '0', 10);
      const cells = X.children(r, 'c');
      let lastCol = 0;
      for (const c of cells) {
        const ref = c.getAttribute('r');
        const pos = ref ? splitRef(ref) : null;
        const col = pos ? pos.col : lastCol + 1;
        const row = pos ? pos.row : rIdx;
        lastCol = col;
        maxCol = Math.max(maxCol, col);
        maxRow = Math.max(maxRow, row);
        cellSeq.push({ c, col, row, rIdx });
      }
    }
    if (maxRow === 0) maxRow = 1;
    if (maxCol === 0) maxCol = 1;

    // 合并
    const mergeMap = new Map();
    const merges = [];
    const mc = X.el(xml, 'mergeCells');
    if (mc) {
      for (const m of X.children(mc, 'mergeCell')) {
        const ref = m.getAttribute('ref');
        const [a, b] = ref.split(':');
        const p1 = splitRef(a), p2 = splitRef(b);
        if (!p1 || !p2) continue;
        merges.push({ r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col });
        mergeMap.set(p1.row + ',' + p1.col, { rs: p2.row - p1.row + 1, cs: p2.col - p1.col + 1 });
        for (let rr = p1.row; rr <= p2.row; rr++)
          for (let cc = p1.col; cc <= p2.col; cc++)
            if (rr !== p1.row || cc !== p1.col) mergeMap.set(rr + ',' + cc, null);
      }
    }

    // 行高
    const rowH = new Map();
    for (const r of rowsXml) {
      const rIdx = parseInt(r.getAttribute('r') || '0', 10);
      if (r.getAttribute('customHeight') === '1' && r.getAttribute('ht')) rowH.set(rIdx, parseFloat(r.getAttribute('ht')) * 4 / 3);
    }

    // 单元格内容
    const cellMap = new Map();
    for (const { c, col, row } of cellSeq) {
      const t = c.getAttribute('t') || 'n';
      const sIdx = parseInt(c.getAttribute('s') || '0', 10);
      const vNode = X.child(c, 'v');
      const fNode = X.child(c, 'f');
      let raw = vNode ? vNode.textContent : null;
      let display = '';
      let isNum = false, num = 0;
      if (t === 's') {
        const idx = parseInt(raw, 10);
        display = sst[idx] != null ? sst[idx] : '';
      } else if (t === 'inlineStr') {
        const isNode = X.child(c, 'is');
        display = isNode ? X.els(isNode, 't').map((x) => x.textContent).join('') : '';
      } else if (t === 'b') {
        display = raw === '1' ? 'TRUE' : 'FALSE';
      } else if (t === 'e') {
        display = raw || '#ERROR';
      } else if (t === 'str') {
        display = raw || '';
      } else if (t === 'd') {
        display = raw ? String(raw).replace('T', ' ').replace(/\.\d+Z?$/, '') : '';
      } else {
        if (raw === null || raw === '') { display = ''; }
        else {
          num = parseFloat(raw);
          isNum = true;
          const xf = styles.xfs[sIdx] || { numFmtId: 0 };
          if (styles.isDate(xf.numFmtId)) {
            const code = styles.fmtCode(xf.numFmtId);
            display = X.formatNumber(num, code || defaultDateFormat(xf.numFmtId));
          } else {
            display = X.formatNumber(num, styles.fmtCode(xf.numFmtId));
          }
        }
      }
      const xf = styles.xfs[sIdx] || { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, align: null };
      cellMap.set(row + ',' + col, { display, xf, isNum, t, formula: !!fNode });
    }

    // 实际渲染范围：裁剪末尾空行空列
    let lastRowWithData = 0, lastColWithData = 0;
    for (const [k] of cellMap) {
      const [r, c] = k.split(',').map(Number);
      if (cellMap.get(k).display !== '') { lastRowWithData = Math.max(lastRowWithData, r); lastColWithData = Math.max(lastColWithData, c); }
    }
    const nRows = Math.max(1, Math.min(lastRowWithData || maxRow, 5000));
    const nCols = Math.max(1, Math.min(lastColWithData || maxCol, 200));

    const colPx = [];
    let totalW = 0;
    for (let i = 1; i <= nCols; i++) {
      const w = colW.has(i) ? colW.get(i) : defaultW;
      const p = Math.max(12, widthPx(w));
      colPx.push(p);
      totalW += p;
    }

    const out = [];
    out.push(`<div class="xls-scroll"><table class="xls" style="width:${totalW + 44}px">`);
    out.push('<colgroup><col style="width:44px">' + colPx.map((w) => `<col style="width:${w}px">`).join('') + '</colgroup>');
    out.push('<thead><tr><th class="xls-corner"></th>' + colPx.map((w, i) => `<th class="xls-ch" style="width:${w}px">${numToCol(i + 1)}</th>`).join('') + '</tr></thead><tbody>');

    let totalH = 24;
    for (let r = 1; r <= nRows; r++) {
      const h = rowH.get(r) || 20;
      totalH += h;
      out.push(`<tr style="height:${h.toFixed(1)}px">`);
      out.push(`<td class="xls-rh" style="height:${h.toFixed(1)}px">${r}</td>`);
      let c = 1;
      while (c <= nCols) {
        const key = r + ',' + c;
        const mg = mergeMap.get(key);
        if (mg === null) { c++; continue; }
        const cell = cellMap.get(key);
        const rs = mg ? mg.rs : 1, cs = mg ? mg.cs : 1;
        const xf = cell ? cell.xf : null;
        const css = [];
        if (xf) {
          const f = styles.fonts[xf.fontId];
          if (f) {
            if (f.name) { const ff = X.fontFamily(f.name); if (ff) css.push('font-family:' + ff); }
            css.push('font-size:' + (f.sz * 4 / 3).toFixed(2) + 'px');
            if (f.b) css.push('font-weight:700');
            if (f.i) css.push('font-style:italic');
            const dec = [];
            if (f.u) dec.push('underline');
            if (f.strike) dec.push('line-through');
            if (dec.length) css.push('text-decoration:' + dec.join(' '));
            if (f.color) css.push('color:' + f.color);
          }
          const fill = styles.fills[xf.fillId];
          if (fill && fill.fg && fill.type && fill.type !== 'none') css.push('background:' + fill.fg);
          const bd = styles.borders[xf.borderId];
          if (bd) {
            for (const k of ['left', 'right', 'top', 'bottom']) {
              const s = bd[k];
              if (!s) { css.push(`border-${k}:1px solid #e4e7ec`); continue; }
              css.push(`border-${k}:${s.style === 'thick' ? 2 : 1}px solid ${s.color || '#9aa4b2'}`);
            }
          }
          if (xf.align) {
            const hm = { center: 'center', right: 'right', justify: 'justify', fill: 'left', distributed: 'justify' }[xf.align.h];
            if (hm) css.push('text-align:' + hm);
            else if (cell && cell.isNum) css.push('text-align:right');
            const vm = { center: 'middle', top: 'top', bottom: 'bottom' }[xf.align.v];
            if (vm) css.push('vertical-align:' + vm);
            if (xf.align.wrap) css.push('white-space:pre-wrap');
            if (xf.align.indent) css.push('padding-left:' + (xf.align.indent * 8 + 4) + 'px');
          } else if (cell && cell.isNum) css.push('text-align:right');
        } else if (cell && cell.isNum) css.push('text-align:right');
        if (!css.some((s) => s.startsWith('text-align')) && cell && cell.isNum) css.push('text-align:right');
        const content = cell ? (cell.isNum && !/\./.test(cell.display) && cell.display !== '' && Math.abs(parseFloat(cell.display)) < 1e15 ? cell.display : X.esc(cell.display)) : '';
        const attrs = [];
        if (cs > 1) attrs.push(`colspan="${cs}"`);
        if (rs > 1) attrs.push(`rowspan="${rs}"`);
        out.push(`<td ${attrs.join(' ')} style="${css.join(';')}">${content}</td>`);
        c += cs;
      }
      out.push('</tr>');
    }
    out.push('</tbody></table></div>');
    return { html: out.join(''), w: totalW + 44, h: totalH, rows: nRows };
  }

  const rendered = [];
  for (const sh of sheets) {
    if (sh.state !== 'visible' && sheets.filter((s) => s.state === 'visible').length > 0) continue;
    const r = renderSheet(sh);
    rendered.push({ name: sh.name, ...r });
  }
  if (!rendered.length) throw new Error('没有可见的工作表');

  return {
    kind: 'xlsx',
    sheets: rendered,
    pageW: rendered[0].w,
    pageH: rendered[0].h,
    flow: true,
    cleanup: () => zip.revoke(),
    text: rendered.map((r) => r.name).join(' / '),
  };
}

function defaultDateFormat(id) {
  const map = {
    14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'yyyy/m/d h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0',
  };
  return map[id] || 'yyyy/m/d';
}
