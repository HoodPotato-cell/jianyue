/* PowerPoint (.pptx) -> HTML 渲染器：幻灯片按原坐标矢量还原 */
import * as X from './oxml.js';
import { OoxmlZip, parseRels, resolvePath, resolveRel } from './zipx.js';

const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DEFAULT_W = 12192000, DEFAULT_H = 6858000;

function relId(node) {
  return node ? (node.getAttributeNS(NS_R, 'id') || node.getAttribute('r:id') || node.getAttribute('id')) : null;
}

function fillCss(node, ctx) {
  if (!node) return null;
  for (const c of X.kids(node)) {
    if (c.localName === 'solidFill') {
      const col = colorOf(c);
      if (col) return col;
    }
    if (c.localName === 'gradFill') {
      const stops = X.els(c, 'gs');
      const cols = [];
      for (const s of stops) { const v = colorOf(X.el(s, 'srgbClr') || X.el(s, 'schemeClr') ? (X.el(s, 'srgbClr') || X.el(s, 'schemeClr')).parentNode : null); const cc = colorOf(c); if (cc) cols.push(cc); break; }
      if (cols.length) return cols[0];
    }
    if (c.localName === 'noFill') return 'transparent';
  }
  return null;
}

function colorOf(fillNode) {
  if (!fillNode) return null;
  for (const c of X.kids(fillNode)) {
    if (c.localName === 'srgbClr') return X.color(X.attr(c, 'val'));
    if (c.localName === 'schemeClr') return X.themeColor(c);
    if (c.localName === 'prstClr') {
      const map = { black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', gray: '#808080', silver: '#c0c0c0' };
      return map[X.attr(c, 'val')] || null;
    }
    if (c.localName === 'sysClr') { const v = X.attr(c, 'lastClr'); return v ? X.color(v) : null; }
  }
  return null;
}

function solidFromProperty(spPr, tag) {
  if (!spPr) return null;
  const f = X.child(spPr, tag);
  if (!f) return null;
  return fillCss(f);
}

function xfrmOf(node) {
  const x = X.el(node, 'xfrm');
  if (!x) return null;
  const off = X.child(x, 'off');
  const ext = X.child(x, 'ext');
  const chOff = X.child(x, 'chOff');
  const chExt = X.child(x, 'chExt');
  return {
    x: off ? X.num(off, 'x') : 0,
    y: off ? X.num(off, 'y') : 0,
    cx: ext ? X.num(ext, 'cx') : 0,
    cy: ext ? X.num(ext, 'cy') : 0,
    rot: X.num(x, 'rot') / 60000,
    chx: chOff ? X.num(chOff, 'x') : 0,
    chy: chOff ? X.num(chOff, 'y') : 0,
    chcx: chExt ? X.num(chExt, 'cx') : 0,
    chcy: chExt ? X.num(chExt, 'cy') : 0,
  };
}

export function renderPptx(bytes, opts = {}) {
  const zip = new OoxmlZip(bytes);
  const presSrc = zip.text('ppt/presentation.xml');
  if (!presSrc) throw new Error('不是有效的 .pptx 文件');
  const pres = X.parseXml(presSrc);
  const presRels = parseRels(zip.text('ppt/_rels/presentation.xml.rels') ? X.parseXml(zip.text('ppt/_rels/presentation.xml.rels')) : null);
  const sldSz = X.el(pres, 'sldSz');
  const slideWEmu = sldSz ? X.num(sldSz, 'cx', DEFAULT_W) : DEFAULT_W;
  const slideHEmu = sldSz ? X.num(sldSz, 'cy', DEFAULT_H) : DEFAULT_H;
  const W = slideWEmu / X.EMU, H = slideHEmu / X.EMU;

  // 主题字体
  let themeMajor = null, themeMinor = null;
  const themePath = findTheme(zip, presRels);
  if (themePath) {
    const t = zip.text(themePath);
    if (t) {
      const tx = X.parseXml(t);
      const mj = X.el(tx, 'majorFont');
      const mn = X.el(tx, 'minorFont');
      const mjEa = mj ? X.el(mj, 'ea') : null;
      const mnEa = mn ? X.el(mn, 'ea') : null;
      if (mjEa) themeMajor = X.attr(mjEa, 'typeface');
      if (mnEa) themeMinor = X.attr(mnEa, 'typeface');
    }
  }

  const slideList = [];
  const idLst = X.el(pres, 'sldIdLst');
  if (idLst) {
    for (const s of X.children(idLst, 'sldId')) {
      const rid = relId(s);
      const rel = presRels[rid];
      if (!rel) continue;
      slideList.push(resolveRel('ppt', rel));
    }
  }

  // 布局/母版占位符位置缓存
  const layoutCache = new Map();
  function loadXml(path) {
    const t = zip.text(path);
    return t ? X.parseXml(t) : null;
  }
  function relsFor(path) {
    const dir = path.split('/').slice(0, -1).join('/');
    const base = path.split('/').pop();
    const p = dir + '/_rels/' + base + '.rels';
    const t = zip.text(p);
    return t ? parseRels(X.parseXml(t)) : {};
  }
  function layoutOf(slidePath) {
    if (layoutCache.has(slidePath)) return layoutCache.get(slidePath);
    let out = { bg: null, phs: new Map() };
    const rl = relsFor(slidePath);
    let layoutPath = null;
    for (const k of Object.keys(rl)) if (/slideLayout$/.test(rl[k].type) || /slideLayouts\//.test(rl[k].target)) layoutPath = resolvePath('ppt/slides', rl[k].target);
    if (!layoutPath) { layoutCache.set(slidePath, out); return out; }
    const lx = loadXml(layoutPath);
    if (lx) {
      out.bg = bgOf(lx);
      collectPhs(lx, out.phs);
      const lrl = relsFor(layoutPath);
      for (const k of Object.keys(lrl)) {
        const tgt = lrl[k].target;
        const p = tgt.includes('slideMaster') ? resolvePath('ppt/slideLayouts', tgt) : null;
        if (p) {
          const mx = loadXml(p);
          if (mx) {
            if (!out.bg) out.bg = bgOf(mx);
            collectPhs(mx, out.phs);
          }
        }
      }
    }
    layoutCache.set(slidePath, out);
    return out;
  }
  function collectPhs(tree, map) {
    const spTree = X.el(tree, 'spTree');
    if (!spTree) return;
    for (const sp of X.children(spTree, 'sp')) {
      const ph = X.el(sp, 'ph');
      if (!ph) continue;
      const idx = X.attr(ph, 'idx') || X.attr(ph, 'type') || 'body';
      if (map.has(idx)) continue;
      const xf = xfrmOf(X.child(sp, 'spPr'));
      if (xf && xf.cx && xf.cy) map.set(idx, xf);
    }
  }
  function bgOf(tree) {
    const bg = X.el(tree, 'bg');
    if (!bg) return null;
    const bgPr = X.child(bg, 'bgPr');
    if (bgPr) {
      const c = solidFromProperty(bgPr, 'solidFill');
      if (c) return { color: c };
      const blip = X.el(bgPr, 'blip');
      if (blip) {
        const rid = X.attr(blip, 'embed');
        return { img: rid };
      }
    }
    const bgRef = X.child(bg, 'bgRef');
    if (bgRef) { const c = colorOf(bgRef); if (c) return { color: c }; }
    return null;
  }

  function imgUrl(slidePath, rid, preRels) {
    const rl = preRels || relsFor(slidePath);
    const r = rl[rid];
    if (!r) return null;
    if (r.external) return r.target;
    const p = resolvePath(slidePath.split('/').slice(0, -1).join('/'), r.target);
    const clean = p.replace(/^ppt\/slides\/\.\.\//, 'ppt/');
    return zip.url(clean) || zip.url(p);
  }

  /* ---------- 母版 / 版式 文本样式继承 ---------- */
  function readLvl(pPrNode) {
    if (!pPrNode) return null;
    const d = {};
    d.algn = X.attr(pPrNode, 'algn') || null;
    const marL = X.attr(pPrNode, 'marL');
    if (marL != null) d.marL = X.num(pPrNode, 'marL', 0) / X.EMU;
    const indent = X.attr(pPrNode, 'indent');
    if (indent != null) d.indent = X.num(pPrNode, 'indent', 0) / X.EMU;
    if (X.child(pPrNode, 'buNone')) d.bullet = null;
    const bc = X.child(pPrNode, 'buChar');
    if (bc) d.bullet = X.attr(bc, 'char') || '•';
    const ba = X.child(pPrNode, 'buAutoNum');
    if (ba) d.autoNum = X.attr(ba, 'type') || 'arabicPeriod';
    const dR = X.child(pPrNode, 'defRPr');
    if (dR) {
      d.sz = X.num(dR, 'sz', null);
      if (X.attr(dR, 'b') === '1') d.b = true;
      if (X.attr(dR, 'i') === '1') d.i = true;
      const sf = X.child(dR, 'solidFill');
      if (sf) { const c = colorOf(sf); if (c) d.color = c; }
      const latin = X.child(dR, 'latin');
      const ea = X.child(dR, 'ea');
      d.face = (ea && X.attr(ea, 'typeface')) || (latin && X.attr(latin, 'typeface')) || null;
    }
    return d;
  }

  function readLstStyle(node) {
    // node: a:lstStyle
    const out = [];
    if (!node) return out;
    for (let i = 1; i <= 9; i++) {
      const lv = X.child(node, 'lvl' + i + 'pPr');
      out[i - 1] = lv ? readLvl(lv) : null;
    }
    return out;
  }

  const styleCache = new Map();
  function masterInfo(slidePath) {
    if (styleCache.has(slidePath)) return styleCache.get(slidePath);
    const info = { title: [], body: [], other: [], layoutPhs: new Map(), masterPhs: new Map() };
    const rels = relsFor(slidePath);
    let layoutPath = null;
    for (const k of Object.keys(rels)) if (/slideLayouts\//.test(rels[k].target)) layoutPath = resolveRel('ppt/slides', rels[k]);
    if (!layoutPath) { styleCache.set(slidePath, info); return info; }
    const lx = loadXml(layoutPath);
    if (lx) {
      info.layoutPhs = phStyles(lx);
      const lrels = relsFor(layoutPath);
      for (const k of Object.keys(lrels)) {
        if (!/slideMasters\//.test(lrels[k].target)) continue;
        const mPath = resolveRel('ppt/slideLayouts', lrels[k]);
        const mx = loadXml(mPath);
        if (!mx) continue;
        info.masterPhs = phStyles(mx);
        const ts = X.el(mx, 'txStyles');
        if (ts) {
          info.title = readLstStyle(X.child(ts, 'titleStyle'));
          info.body = readLstStyle(X.child(ts, 'bodyStyle'));
          info.other = readLstStyle(X.child(ts, 'otherStyle'));
        }
        break;
      }
    }
    styleCache.set(slidePath, info);
    return info;
  }

  function phStyles(tree) {
    const map = new Map();
    const spTree = X.el(tree, 'spTree');
    if (!spTree) return map;
    for (const sp of X.children(spTree, 'sp')) {
      const ph = X.el(sp, 'ph');
      if (!ph) continue;
      const idx = X.attr(ph, 'idx');
      const type = X.attr(ph, 'type') || 'body';
      const txBody = X.child(sp, 'txBody');
      const bodyPr = txBody ? X.child(txBody, 'bodyPr') : null;
      const entry = {
        anchor: bodyPr ? X.attr(bodyPr, 'anchor') : null,
        lst: txBody ? readLstStyle(X.child(txBody, 'lstStyle')) : [],
        xfrm: xfrmOf(X.child(sp, 'spPr')),
      };
      if (idx != null && !map.has('i' + idx)) map.set('i' + idx, entry);
      if (!map.has('t' + type)) map.set('t' + type, entry);
    }
    return map;
  }

  function styleFor(slidePath, phType, phIdx) {
    const info = masterInfo(slidePath);
    const isTitle = phType === 'title' || phType === 'ctrTitle';
    const base = isTitle ? info.title : (phType === 'body' || phType === 'obj' || phType === 'subTitle' ? info.body : (info.other.length ? info.other : info.body));
    const lay = (phIdx != null && info.layoutPhs.get('i' + phIdx)) || info.layoutPhs.get('t' + phType) || null;
    const mas = (phIdx != null && info.masterPhs.get('i' + phIdx)) || info.masterPhs.get('t' + phType) || null;
    const lvls = [];
    for (let i = 0; i < 9; i++) {
      const merged = { ...(base[i] || {}) };
      const m2 = mas && mas.lst[i]; if (m2) Object.assign(merged, m2);
      const l2 = lay && lay.lst[i]; if (l2) Object.assign(merged, l2);
      lvls.push(merged);
    }
    return { lvls, anchor: (lay && lay.anchor) || (mas && mas.anchor) || null };
  }

  function renderTextBody(txBody, styleDef, ctx) {
    if (!txBody) return '';
    const bodyPr = X.child(txBody, 'bodyPr');
    const anchor = (bodyPr && X.attr(bodyPr, 'anchor')) || (styleDef && styleDef.anchor) || 't';
    const ins = bodyPr ? {
      l: X.num(bodyPr, 'lIns', 91440) / X.EMU,
      r: X.num(bodyPr, 'rIns', 91440) / X.EMU,
      t: X.num(bodyPr, 'tIns', 45720) / X.EMU,
      b: X.num(bodyPr, 'bIns', 45720) / X.EMU,
    } : { l: 9.6, r: 9.6, t: 4.8, b: 4.8 };
    let scale = 1;
    const noAuto = X.child(bodyPr, 'normAutofit');
    if (noAuto) {
      const fs = X.num(noAuto, 'fontScale', 100000);
      if (fs && fs < 100000) scale = fs / 100000;
    }
    const ownLst = readLstStyle(X.child(txBody, 'lstStyle'));
    const ownBodyPr = bodyPr;
    const paras = X.children(txBody, 'p');
    const justify = { t: 'flex-start', ctr: 'center', b: 'flex-end', just: 'space-between', dist: 'space-between' }[anchor] || 'flex-start';

    const numCounters = new Map();
    let html = '';
    for (const p of paras) {
      const pPr = X.child(p, 'pPr');
      const lvl = pPr ? X.num(pPr, 'lvl', 0) : 0;
      const ownLvl = ownLst[lvl] || null;
      const stLvl = (styleDef && styleDef.lvls[lvl]) || {};
      const lvlDef = { ...stLvl, ...(ownLvl || {}) };
      const ownRead = readLvl(pPr) || {};
      const merged = { ...lvlDef, ...ownRead };

      let bullet = merged.bullet;
      if (merged.autoNum) {
        const n = (numCounters.get(lvl) || 0) + 1;
        numCounters.set(lvl, n);
        bullet = autoNumText(merged.autoNum, n);
      }
      const marL = merged.marL || 0;
      const indent = merged.indent || 0;

      const lh = pPr ? lineSpacing(pPr) : null;
      const before = pPr ? spacingPts(X.child(pPr, 'spcBef')) : 0;
      const after = pPr ? spacingPts(X.child(pPr, 'spcAft')) : 0;
      const algn = (pPr && X.attr(pPr, 'algn')) || merged.algn || 'l';

      const pCss = [];
      pCss.push('text-align:' + ({ l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[algn] || 'left'));
      if (marL) pCss.push('margin-left:' + marL.toFixed(1) + 'px');
      if (indent) pCss.push('text-indent:' + indent.toFixed(1) + 'px');
      if (before) pCss.push('margin-top:' + before.toFixed(1) + 'px');
      if (after) pCss.push('margin-bottom:' + after.toFixed(1) + 'px');
      if (lh) pCss.push('line-height:' + lh);

      const runs = [];
      const defRPr = pPr ? X.child(pPr, 'defRPr') : null;
      for (const c of X.kids(p)) {
        if (c.localName === 'r' || c.localName === 'fld') {
          const rPr = X.child(c, 'rPr');
          const t = X.child(c, 't');
          runs.push(runHtml(t ? t.textContent : '', rPr, defRPr, merged, scale, ctx));
        } else if (c.localName === 'br') {
          runs.push('<br>');
        }
      }
      let inner = runs.join('');
      if (bullet) {
        const indWidth = indent < 0 ? -indent : 14;
        inner = `<span style="display:inline-block;width:${indWidth.toFixed(1)}px">${X.esc(bullet)}</span>` + inner;
        if (!marL) pCss.push('margin-left:0');
      }
      if (!inner) inner = '<br>';
      html += `<div class="ppt-p" style="${pCss.join(';')}">${inner}</div>`;
    }
    if (scale < 1) html = `<div style="zoom:${scale.toFixed(3)}">${html}</div>`;
    return `<div class="ppt-tx" style="position:absolute;left:0;top:0;width:100%;height:100%;display:flex;flex-direction:column;justify-content:${justify};box-sizing:border-box;padding:${ins.t}px ${ins.r}px ${ins.b}px ${ins.l}px;overflow:hidden">${html}</div>`;
  }

  function autoNumText(type, n) {
    const roman = (v) => { const R = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; for (const [a, b] of R) while (v >= a) { s += b; v -= a; } return s; };
    const alpha = (v) => { let s = ''; while (v > 0) { const r = (v - 1) % 26; s = String.fromCharCode(97 + r) + s; v = Math.floor((v - 1) / 26); } return s; };
    if (/roman/i.test(type)) return roman(n) + (type.endsWith('Paren') ? ')' : '.');
    if (/alpha/i.test(type)) return (/Uc/.test(type) ? alpha(n).toUpperCase() : alpha(n)) + (type.endsWith('Paren') ? ')' : '.');
    if (/chinese/i.test(type)) return ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'][Math.min(n, 10)] + '、';
    if (type.endsWith('Paren')) return n + ')';
    if (type.endsWith('Period')) return n + '.';
    return n + '.';
  }

  function lineSpacing(pPr) {
    const ls = X.child(pPr, 'lnSpc');
    if (!ls) return null;
    const pct = X.child(ls, 'spcPct');
    if (pct) return (X.num(pct, 'val', 100000) / 100000).toFixed(3);
    const pts = X.child(ls, 'spcPts');
    if (pts) return (X.num(pts, 'val', 1200) / 100 * X.PT).toFixed(1) + 'px';
    return null;
  }
  function spacingPts(node) {
    if (!node) return 0;
    const pct = X.child(node, 'spcPct');
    if (pct) return 0;
    const pts = X.child(node, 'spcPts');
    if (pts) return (X.num(pts, 'val', 0) / 100) * X.PT;
    return 0;
  }

  function runHtml(text, rPr, defRPr, styleLvl, scale, ctx) {
    if (!text) return '';
    const eff = styleLvl || {};
    const st = [];
    let szPt = null;
    if (rPr && X.attr(rPr, 'sz') != null) szPt = X.num(rPr, 'sz') / 100;
    else if (defRPr && X.attr(defRPr, 'sz') != null) szPt = X.num(defRPr, 'sz') / 100;
    else if (eff.sz != null) szPt = eff.sz / 100;
    if (szPt == null) szPt = 18;
    st.push('font-size:' + (szPt * scale * X.PT).toFixed(2) + 'px');

    let bold = false, ital = false, color = null, face = null, baseline = null, underline = false;
    if (rPr) {
      if (X.attr(rPr, 'b') === '1') bold = true;
      if (X.attr(rPr, 'i') === '1') ital = true;
      if (X.attr(rPr, 'u') && X.attr(rPr, 'u') !== 'none') underline = true;
      const sf = X.child(rPr, 'solidFill');
      if (sf) color = colorOf(sf);
      const latin = X.child(rPr, 'latin');
      const ea = X.child(rPr, 'ea');
      face = (ea && X.attr(ea, 'typeface')) || (latin && X.attr(latin, 'typeface')) || null;
      baseline = X.attr(rPr, 'baseline');
    }
    if (bold) st.push('font-weight:700');
    if (ital || eff.i) st.push('font-style:italic');
    if (underline) st.push('text-decoration:underline');
    if (color) st.push('color:' + color);
    else if (eff.color) st.push('color:' + eff.color);
    if (!color && !eff.color && X.attr(rPr || {}, 'b') === '1') st.push('color:#000');
    if (!face) face = eff.face;
    if (face && !face.startsWith('+')) { const ff = X.fontFamily(face); if (ff) st.push('font-family:' + ff); }
    if (baseline && baseline !== '0') st.push('vertical-align:' + (parseFloat(baseline) > 0 ? 'super' : 'sub'), 'font-size:0.72em');
    return `<span style="${st.join(';')}">${X.esc(text)}</span>`;
  }

  function renderSp(sp, slidePath, rels, ctx) {
    const spPr = X.child(sp, 'spPr');
    let xf = xfrmOf(spPr);
    const phNode = X.el(sp, 'ph');
    const phType = phNode ? (X.attr(phNode, 'type') || 'body') : null;
    const phIdx = phNode ? X.attr(phNode, 'idx') : null;
    if (!xf || !xf.cx || !xf.cy) {
      const info = masterInfo(slidePath);
      const pick = (m) => {
        let e = null;
        if (phIdx != null) e = m.get('i' + phIdx);
        if (!e && phType) e = m.get('t' + phType);
        return (e && e.xfrm && e.xfrm.cx) ? e : null;
      };
      const lph = pick(info.layoutPhs) || pick(info.masterPhs);
      if (lph) xf = lph.xfrm;
    }
    if (!xf || !xf.cx) return '';
    const st = ['position:absolute'];
    st.push(`left:${(xf.x / X.EMU).toFixed(2)}px`);
    st.push(`top:${(xf.y / X.EMU).toFixed(2)}px`);
    st.push(`width:${(xf.cx / X.EMU).toFixed(2)}px`);
    st.push(`height:${(xf.cy / X.EMU).toFixed(2)}px`);
    const fill = solidFromProperty(spPr, 'solidFill');
    if (fill) st.push('background:' + fill);
    if (spPr) {
      const ln = X.child(spPr, 'ln');
      if (ln) {
        const lc = colorOf(X.el(ln, 'solidFill'));
        const lw = X.num(ln, 'w', 0) / X.EMU;
        if (lc && lw) st.push(`border:${lw.toFixed(2)}px solid ${lc}`);
      }
      const geo = X.child(spPr, 'prstGeom');
      if (geo) {
        const prst = X.attr(geo, 'prst');
        if (prst === 'ellipse') st.push('border-radius:50%');
        else if (prst === 'roundRect') st.push('border-radius:' + (Math.min(xf.cx, xf.cy) / X.EMU * 0.16).toFixed(1) + 'px');
      }
    }
    if (xf.rot) st.push(`transform:rotate(${xf.rot.toFixed(2)}deg)`);
    const styleDef = phNode ? styleFor(slidePath, phType, phIdx) : null;
    const body = renderTextBody(X.child(sp, 'txBody'), styleDef, ctx);
    return `<div class="ppt-shape" style="${st.join(';')}">${body}</div>`;
  }

  function renderPic(pic, slidePath, rels, ctx) {
    const spPr = X.child(pic, 'spPr');
    let xf = xfrmOf(spPr);
    if (!xf || !xf.cx) {
      const ph = X.el(pic, 'ph');
      if (ph) {
        const info = masterInfo(slidePath);
        const idx = X.attr(ph, 'idx');
        const ty = X.attr(ph, 'type') || 'body';
        const pick = (m) => { const e = (idx != null ? m.get('i' + idx) : null) || m.get('t' + ty); return (e && e.xfrm && e.xfrm.cx) ? e : null; };
        const lph = pick(info.layoutPhs) || pick(info.masterPhs);
        if (lph) xf = lph.xfrm;
      }
    }
    if (!xf) return '';
    const blip = X.el(pic, 'blip');
    const rid = blip ? X.attr(blip, 'embed') : null;
    const src = rid ? imgUrl(slidePath, rid, rels) : null;
    const st = ['position:absolute',
      `left:${(xf.x / X.EMU).toFixed(2)}px`, `top:${(xf.y / X.EMU).toFixed(2)}px`,
      `width:${(xf.cx / X.EMU).toFixed(2)}px`, `height:${(xf.cy / X.EMU).toFixed(2)}px`];
    if (xf.rot) st.push(`transform:rotate(${xf.rot.toFixed(2)}deg)`);
    if (!src) return `<div style="${st.join(';')}"></div>`;
    return `<div style="${st.join(';')}"><img src="${src}" style="width:100%;height:100%;object-fit:contain;display:block" alt=""></div>`;
  }

  function renderTable(frame, slidePath, rels, ctx) {
    const spPr = X.child(frame, 'spPr');
    const xf = xfrmOf(spPr);
    const tbl = X.el(frame, 'tbl');
    if (!tbl) return '';
    const st = ['position:absolute'];
    if (xf) {
      st.push(`left:${(xf.x / X.EMU).toFixed(2)}px`, `top:${(xf.y / X.EMU).toFixed(2)}px`,
        `width:${(xf.cx / X.EMU).toFixed(2)}px`, `height:${(xf.cy / X.EMU).toFixed(2)}px`);
    }
    const gridCols = X.children(X.el(tbl, 'tblGrid'), 'gridCol');
    const totalW = gridCols.reduce((a, c) => a + X.num(c, 'w'), 0) || 1;
    const pct = gridCols.map((c) => (X.num(c, 'w') / totalW) * 100);
    let html = `<table style="border-collapse:collapse;table-layout:fixed;width:100%;height:100%;font-size:${(14 * X.PT).toFixed(1)}px">`;
    if (pct.length) html += '<colgroup>' + pct.map((p) => `<col style="width:${p.toFixed(3)}%">`).join('') + '</colgroup>';
    for (const tr of X.children(tbl, 'tr')) {
      html += '<tr>';
      for (const tc of X.children(tr, 'tc')) {
        const tcPr = X.child(tc, 'tcPr');
        const css = ['border:1px solid #c9ced6', 'padding:4px 6px', 'vertical-align:middle'];
        const f = tcPr ? fillCss(X.child(tcPr, 'solidFill') ? tcPr : null) : null;
        const fill = tcPr ? solidFromProperty(tcPr, 'solidFill') : null;
        if (fill) css.push('background:' + fill);
        const gs = tcPr ? X.num(tcPr, 'gridSpan', 1) : 1;
        html += `<td${gs > 1 ? ` colspan="${gs}"` : ''} style="${css.join(';')}">${renderTextBody(X.child(tc, 'txBody'), null, ctx)}</td>`;
      }
      html += '</tr>';
    }
    html += '</table>';
    return `<div style="${st.join(';')}">${html}</div>`;
  }

  function renderGroup(grp, slidePath, rels, ctx) {
    const spPr = X.child(grp, 'grpSpPr');
    const xf = xfrmOf(spPr);
    if (!xf) return '';
    const sx = xf.chcx ? xf.cx / xf.chcx : 1;
    const sy = xf.chcy ? xf.cy / xf.chcy : 1;
    const tx = xf.x / X.EMU - (xf.chx * sx) / X.EMU;
    const ty = xf.y / X.EMU - (xf.chy * sy) / X.EMU;
    const inner = renderShapeTree(grp, slidePath, rels, ctx);
    return `<div style="position:absolute;left:0;top:0;width:100%;height:100%;transform-origin:0 0;transform:translate(${tx.toFixed(2)}px,${ty.toFixed(2)}px) scale(${sx},${sy})">${inner}</div>`;
  }

  function renderShapeTree(tree, slidePath, rels, ctx) {
    const out = [];
    for (const c of X.kids(tree)) {
      switch (c.localName) {
        case 'sp': out.push(renderSp(c, slidePath, rels, ctx)); break;
        case 'pic': out.push(renderPic(c, slidePath, rels, ctx)); break;
        case 'graphicFrame': out.push(renderTable(c, slidePath, rels, ctx)); break;
        case 'grpSp': out.push(renderGroup(c, slidePath, rels, ctx)); break;
        case 'cxnSp': out.push(renderSp(c, slidePath, rels, ctx)); break;
        default: break;
      }
    }
    return out.join('');
  }

  const pages = [];
  slideList.forEach((path, idx) => {
    const xml = loadXml(path);
    if (!xml) return;
    const rels = relsFor(path);
    const cSld = X.el(xml, 'cSld');
    const layout = layoutOf(path);
    let bgCss = '#ffffff';
    let bgImg = null;
    const ownBg = bgOf(xml);
    const bg = ownBg || layout.bg;
    if (bg) {
      if (bg.color) bgCss = bg.color;
      if (bg.img) { const u = bg.img; bgImg = imgUrl(path, u, rels); }
    }
    const spTree = cSld ? X.el(cSld, 'spTree') : X.el(xml, 'spTree');
    const inner = spTree ? renderShapeTree(spTree, path, rels, {}) : '';
    const style = [`width:${W}px`, `height:${H}px`, 'position:relative', 'background:' + bgCss, 'overflow:hidden'];
    if (bgImg) style.push(`background-image:url('${bgImg}')`, 'background-size:cover');
    pages.push({
      idx,
      w: W, h: H,
      html: `<div class="ppt-canvas" style="${style.join(';')}">${inner}</div>`,
    });
  });

  if (!pages.length) throw new Error('幻灯片为空');

  return {
    kind: 'pptx',
    pages,
    pageW: W,
    pageH: H,
    flow: false,
    cleanup: () => zip.revoke(),
    text: '',
  };
}

function findTheme(zip, presRels) {
  for (const k of Object.keys(presRels)) {
    const r = presRels[k];
    if (r.type && /theme$/.test(r.type)) return resolvePath('ppt', r.target);
  }
  return 'ppt/theme/theme1.xml';
}
