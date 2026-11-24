/* 简阅 · 应用主逻辑 */
import { PdfViewer } from './pdfview.js';
import { renderDocx } from './docx.js';
import { renderXlsx } from './xlsx.js';
import { renderPptx } from './pptx.js';
import { renderMarkdown } from './markdown.js';
import { parseDoc, docHtml } from './doc97.js';
import { parseXls } from './xls97.js';
import { parsePpt, fitPptText } from './ppt97.js';

const TOKEN = new URLSearchParams(location.search).get('t') || '';
const $ = (id) => document.getElementById(id);

const el = {
  shelf: $('shelf'), reader: $('reader'),
  list: $('list'), state: $('state'), q: $('q'), qClear: $('qClear'),
  chips: $('chips'), btnRefresh: $('btnRefresh'), brandSub: $('brandSub'),
  permBar: $('permBar'), btnPerm: $('btnPerm'),
  stage: $('stage'), stageB: $('stageB'), panes: $('panes'), paneA: $('paneA'), paneB: $('paneB'),
  loading: $('loading'), loadingTxt: $('loadingTxt'),
  pickerSide: $('pickerSide'), pickerMask: $('pickerMask'), psList: $('psList'),
  psClose: $('psClose'), psOrient: $('psOrient'), btnSplit: $('btnSplit'), btnTheme: $('btnTheme'),
  rName: $('rName'), rMeta: $('rMeta'), btnBack: $('btnBack'), btnMore: $('btnMore'),
  zIn: $('zIn'), zOut: $('zOut'), zVal: $('zVal'), btnFit: $('btnFit'),
  pager: $('pager'), pPrev: $('pPrev'), pNext: $('pNext'), pNow: $('pNow'),
  sheet: $('sheet'), sheetMask: $('sheetMask'), toast: $('toast'),
  jumpBar: $('jumpBar'), jumpText: $('jumpText'), jumpGo: $('jumpGo'), jumpClose: $('jumpClose'),
  sheetTitle: $('sheetTitle'), sheetBody: $('sheetBody'),
  dlg: $('dlg'), dlgMask: $('dlgMask'), dlgTitle: $('dlgTitle'),
  dlgText: $('dlgText'), dlgOk: $('dlgOk'), dlgCancel: $('dlgCancel'), dlgInput: $('dlgInput'),
  topbar: $('topbar'), btnSelect: $('btnSelect'), btnAdd: $('btnAdd'), editbar: $('editbar'),
  btnExitEdit: $('btnExitEdit'), ebTitle: $('ebTitle'), btnSelAll: $('btnSelAll'),
  actionbar: $('actionbar'), btnMerge: $('btnMerge'), btnDelSel: $('btnDelSel'),
  shelves: $('shelves'), colbar: $('colbar'), colTitle: $('colTitle'),
  btnColBack: $('btnColBack'), btnColMenu: $('btnColMenu'),
};

const MANUAL_EXTS = new Set(['md', 'markdown']);

/* 分屏：a 为主面板，b 为分屏面板 */
const PANES = {
  a: { stage: null, view: null, entry: null },
  b: { stage: null, view: null, entry: null },
};
let ACT = 'a';
const SPLIT = { on: false, vertical: false };
const curStage = () => PANES[ACT].stage;
const curView = () => PANES[ACT].view;
function setView(v) {
  PANES[ACT].view = v;
  if (ACT === 'a') S.view = v; else S.viewB = v;
}
const curEntry = () => PANES[ACT].entry;
function setEntry(f) {
  PANES[ACT].entry = f;
  if (ACT === 'a') S.entry = f; else S.entryB = f;
}

/* 主题 */
function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  if (el.btnTheme) el.btnTheme.setAttribute('aria-label', t === 'dark' ? '切换到浅色' : '切换到深色');
}
function initTheme() {
  let t = '';
  try { t = localStorage.getItem('jy_theme') || ''; } catch (e) {}
  if (!t) t = (window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
  applyTheme(t);
}
function toggleTheme() {
  const now = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  applyTheme(now);
  savePrefs();
}

const S = {
  files: [], filter: 'all', q: '', scanning: false, scannedAt: 0, manual: new Set(),
  entry: null, view: null, zoom: 1, fit: 1, kind: null, busy: false,
  collections: [], collectionId: null, editing: false, selected: new Set(), fileMap: new Map(),
};

const TYPE = {
  pdf: { cls: 'pdf', label: 'PDF', name: 'PDF' },
  doc: { cls: 'doc', label: 'DOC', name: 'Word' },
  docx: { cls: 'doc', label: 'DOC', name: 'Word' },
  xls: { cls: 'xls', label: 'XLS', name: 'Excel' },
  xlsx: { cls: 'xls', label: 'XLS', name: 'Excel' },
  ppt: { cls: 'ppt', label: 'PPT', name: 'PPT' },
  pptx: { cls: 'ppt', label: 'PPT', name: 'PPT' },
  md: { cls: 'md', label: 'MD', name: 'Markdown' },
};
const groupOf = (e) => (e === 'pdf' ? 'pdf' : (e === 'doc' || e === 'docx') ? 'doc' : (e === 'xls' || e === 'xlsx') ? 'xls' : (e === 'md' || e === 'markdown') ? 'md' : 'ppt');

/* ------------------------------------------------------------------ 工具 */
/* WebView 每次启动端口随机，localStorage 按源隔离不可靠；
   偏好与手动添加改存到本地服务的书库里（Java 侧持久化） */
const PREFS_COL = '内部·设置（勿改）';
const POS_PREFIX = 'pos:';
const S_POS = new Map();   // 文件路径 → 阅读进度（0..1）

async function loadPrefs() {
  try {
    const d = await (await fetch('/api/collections?t=' + encodeURIComponent(TOKEN), { cache: 'no-store' })).json();
    const c = (d.collections || []).find((x) => x.name === PREFS_COL);
    const items = c && Array.isArray(c.items) ? c.items : [];
    S.manual = new Set(items.filter((i) => i.startsWith('md:')).map((i) => i.slice(3)));
    for (const it of items) {
      if (!it.startsWith(POS_PREFIX)) continue;
      const body = it.slice(POS_PREFIX.length);
      const eq = body.lastIndexOf('=');
      if (eq <= 0) continue;
      S_POS.set(decodeURIComponent(body.slice(0, eq)), Number(body.slice(eq + 1)) || 0);
    }
    const th = items.find((i) => i.startsWith('theme:'));
    if (th === 'theme:dark' || th === 'theme:light') applyTheme(th.slice(6));
  } catch (e) {
    S.manual = S.manual || new Set();
  }
}

async function savePrefs() {
  try {
    const items = [...S.manual].map((p) => 'md:' + p);
    const th = document.documentElement.dataset.theme;
    if (th) items.push('theme:' + th);
    let n = 0;
    for (const [path, r] of S_POS) {
      if (n++ >= 300) break;
      items.push(POS_PREFIX + encodeURIComponent(path) + '=' + r.toFixed(4));
    }
    const d = await (await fetch('/api/collections?t=' + encodeURIComponent(TOKEN), { cache: 'no-store' })).json();
    const c = (d.collections || []).find((x) => x.name === PREFS_COL);
    if (c) {
      await fetch('/api/collection/delete?t=' + encodeURIComponent(TOKEN) + '&id=' + encodeURIComponent(c.id), { method: 'POST' });
    }
    if (items.length) {
      await fetch('/api/collection/create?t=' + encodeURIComponent(TOKEN), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: PREFS_COL, items }),
      });
    }
  } catch (e) { /* 静默 */ }
}

function fmtSize(n) {
  if (!n) return '—';
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(0) + ' KB';
  if (n < 1024 * 1024 * 1024) return (n / 1048576).toFixed(1) + ' MB';
  return (n / 1073741824).toFixed(2) + ' GB';
}
function fmtDate(t) {
  if (!t) return '';
  const d = new Date(t);
  const now = new Date();
  const same = d.toDateString() === now.toDateString();
  const p = (x) => String(x).padStart(2, '0');
  if (same) return '今天 ' + p(d.getHours()) + ':' + p(d.getMinutes());
  const y = d.getFullYear() === now.getFullYear();
  return (y ? '' : d.getFullYear() + '/') + (d.getMonth() + 1) + '/' + d.getDate();
}
function toast(msg, ms = 2200) {
  el.toast.textContent = msg;
  el.toast.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { el.toast.hidden = true; }, ms);
}
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/* ------------------------------------------------------------------ 数据 */
async function apiFiles() {
  const r = await fetch('/api/files?t=' + encodeURIComponent(TOKEN), { cache: 'no-store' });
  return r.json();
}
async function apiRescan() {
  try { await fetch('/api/rescan?t=' + encodeURIComponent(TOKEN)); } catch (e) {}
}
const fileUrl = (p) => '/api/file?p=' + encodeURIComponent(p) + '&t=' + encodeURIComponent(TOKEN);

/* ------------------------------------------------------------------ 书库 */
function filtered() {
  const q = S.q.trim().toLowerCase();
  const col = S.collections.find((c) => c.id === S.collectionId);
  const inCol = col ? new Set(col.items) : null;
  return S.files.filter((f) => {
    if (MANUAL_EXTS.has(f.e) && !S.manual.has(f.p)) return false;
    if (inCol && !inCol.has(f.p)) return false;
    if (S.filter !== 'all' && groupOf(f.e) !== S.filter) return false;
    if (q && f.n.toLowerCase().indexOf(q) < 0 && (f.d || '').toLowerCase().indexOf(q) < 0) return false;
    return true;
  });
}

function renderList() {
  const list = filtered();
  if (!list.length) {
    el.list.innerHTML = '';
    el.state.classList.remove('hide');
    const hiddenN = S.files.filter((f) => MANUAL_EXTS.has(f.e) && !S.manual.has(f.p)).length;
    el.state.innerHTML = S.scanning
      ? '<div class="spinner"></div><p>正在扫描本机文档…</p>'
      : (!S.collectionId && !S.q && S.filter === 'all' && hiddenN && !S.manual.size)
        ? '<p style="font-size:15px;color:#5a6270">还没有添加文档</p><p style="font-size:12.5px">点右上角「＋」可以手动添加 Markdown 文件</p>'
      : (S.collectionId ? '<p style="font-size:15px;color:#5a6270">这个书库里还没有文档</p><p style="font-size:12.5px">点「选择」勾选文档后可以合并进来</p>'
        : S.files.length ? '<p style="font-size:15px;color:#5a6270">没有匹配的文档</p><p style="font-size:12.5px">试试其它关键词</p>'
        : '<p style="font-size:15px;color:#5a6270">还没有找到文档</p><p style="font-size:12.5px">把 PDF / Word / Excel / PPT 放到手机里，<br>或在下方点击刷新重新扫描</p>');
    return;
  }
  el.state.classList.add('hide');
  const frag = document.createDocumentFragment();
  for (const f of list) {
    const t = TYPE[f.e] || TYPE.pdf;
    const d = document.createElement('button');
    d.type = 'button';
    d.className = 'item';
    d.dataset.p = f.p;
    if (S.selected.has(f.p)) d.classList.add('on');
    d.innerHTML =
      `<span class="sel"></span>` +
      `<div class="tile ${t.cls}">${t.label}</div>` +
      `<div class="item-main">` +
      `<div class="item-name">${esc(f.n.replace(/\.[^.]+$/, ''))}</div>` +
      `<div class="item-sub"><span>${fmtSize(f.s)}</span><i>·</i><span>${fmtDate(f.t)}</span><i>·</i><span>${esc(shortDir(f.d))}</span></div>` +
      `</div>`;
    frag.appendChild(d);
  }
  el.list.innerHTML = '';
  el.list.appendChild(frag);
  el.btnSelect.hidden = !S.files.length || S.editing;
}

function shortDir(d) {
  if (!d) return '';
  const parts = d.split('/').filter(Boolean);
  const tail = parts.slice(-2).join('/');
  return tail === 'emulated/0' ? '内部存储' : tail;
}

async function loadFiles(showSpin, pollRounds) {
  if (showSpin) {
    el.state.classList.remove('hide');
    el.state.innerHTML = '<div class="spinner"></div><p>正在扫描本机文档…</p>';
  }
  try {
    const d = await apiFiles();
    S.files = d.files || [];
    S.fileMap = new Map(S.files.map((f) => [f.p, f]));
    S.scanning = !!d.scanning;
    S.scannedAt = d.scannedAt || 0;
    if (!S.scanning && S.files.length) {
      let changed = false;
      for (const fp of [...S.manual]) if (!S.fileMap.has(fp)) { S.manual.delete(fp); changed = true; }
      if (changed) savePrefs();
    }
    const visible = S.files.filter((f) => !MANUAL_EXTS.has(f.e) || S.manual.has(f.p)).length;
    el.brandSub.textContent = visible ? `${visible} 份文档 · 自动收录` : '文档阅读器';
  if (S.files.length && !S.hinted) {
    S.hinted = true;
    let seen = false;
    try { seen = localStorage.getItem('jy_hint_longpress') === '1'; } catch (e) {}
    if (!seen) {
      setTimeout(() => {
        toast('提示：长按文档可以删除或查看更多操作', 3600);
        try { localStorage.setItem('jy_hint_longpress', '1'); } catch (e) {}
      }, 900);
    }
  }
  } catch (e) {
    el.brandSub.textContent = '无法连接服务';
  }
  renderList();
  renderShelves();
  if (S.scanning) setTimeout(() => { if (!S.busy) loadFiles(false, 3); }, 900);
  else if (pollRounds > 0) setTimeout(() => { if (!S.busy) loadFiles(false, pollRounds - 1); }, 1300);
}

/* ------------------------------------------------------------------ 视图层 */
class OfficeView {
  constructor(stage, doc, hostEl, width) {
    this.stage = stage;
    this.doc = doc;
    this.host = hostEl;
    this.baseW = width;
    this.zoom = 1;
    this.type = doc.kind;
    this.total = doc.kind === 'xlsx' ? 1 : (doc.pages ? doc.pages.length : doc.pageCount || 1);
  }
  get contentEl() { return this.host; }
  get layoutZoomed() { return true; }
  fitZoom() {
    const avail = this.stage.clientWidth - 20;
    return Math.max(0.15, avail / this.baseW);
  }
  setZoom(z) {
    this.zoom = Math.max(0.2, Math.min(z, 8));
    this.host.style.zoom = this.zoom;
    return Promise.resolve();
  }
  currentPage() {
    if (this.type === 'xlsx') return 1;
    if (this.type === 'pptx') {
      const pages = this.host.querySelectorAll('.ppt-wrap');
      const top = this.stage.scrollTop + 60;
      for (let i = 0; i < pages.length; i++) {
        if (pages[i].offsetTop + pages[i].offsetHeight > top) return i + 1;
      }
      return pages.length;
    }
    // docx：按滚动比例估算
    const h = this.stage.scrollHeight - this.stage.clientHeight;
    const r = h > 0 ? this.stage.scrollTop / h : 0;
    return Math.min(this.total, Math.max(1, Math.round(r * (this.total - 1)) + 1));
  }
  gotoPage(n) {
    if (this.type === 'xlsx') return;
    if (this.type === 'pptx') {
      const pages = this.host.querySelectorAll('.ppt-wrap');
      const p = pages[Math.max(0, Math.min(n - 1, pages.length - 1))];
      if (p) this.stage.scrollTop = p.offsetTop - 14;
    } else {
      const h = this.stage.scrollHeight - this.stage.clientHeight;
      this.stage.scrollTop = (n - 1) / Math.max(1, this.total - 1) * h;
    }
  }
  destroy() {
    this.stage.innerHTML = '';
    if (this.doc && this.doc.cleanup) { try { this.doc.cleanup(); } catch (e) {} }
  }
}

class PdfDocView {
  constructor(stage, url, onState) {
    this.stage = stage;
    this.viewer = new PdfViewer(stage, onState);
    this.url = url;
    this.total = 0;
    this.zoom = 1;
  }
  async load() {
    await this.viewer.load(this.url);
    this.total = this.viewer.pages.length;
  }
  get contentEl() { return this.viewer.wrap; }
  get layoutZoomed() { return false; }
  fitZoom() { return this.viewer.fitWidth(24); }
  setZoom(z) {
    this.zoom = z;
    return this.viewer.setZoom(z);
  }
  currentPage() { return this.viewer.currentPage(); }
  gotoPage(n) { this.viewer.gotoPage(n); }
  destroy() { this.viewer.destroy(); }
}

/* ------------------------------------------------------------------ 阅读器 */
function showReader(on) {
  el.shelf.classList.toggle('active', !on);
  el.reader.classList.toggle('active', on);
}

function assignView(k, v) {
  PANES[k].view = v;
  if (k === 'a') S.view = v; else S.viewB = v;
}
function assignEntry(k, f) {
  PANES[k].entry = f;
  if (k === 'a') S.entry = f; else S.entryB = f;
}


/* ---------------- 阅读进度 ---------------- */
let posTimer = 0;

function recordPos(k) {
  const pane = PANES[k];
  if (!pane || !pane.entry || !pane.stage) return;
  const st = pane.stage;
  const denom = Math.max(1, st.scrollHeight - st.clientHeight);
  const r = Math.min(1, Math.max(0, st.scrollTop / denom));
  S_POS.set(pane.entry.p, r);
}

function schedulePosSave(k) {
  clearTimeout(posTimer);
  posTimer = setTimeout(() => {
    recordPos(k);
    savePrefs();
  }, 1200);
}

let jumpTarget = null;

function applyJump(k, r) {
  const pane = PANES[k];
  if (!pane || !pane.stage) return;
  const st = pane.stage;
  const go = () => {
    const denom = Math.max(0, st.scrollHeight - st.clientHeight);
    if (denom > 0) st.scrollTop = r * denom;
  };
  requestAnimationFrame(go);
  setTimeout(go, 260);
}

function hideJumpBar() {
  jumpTarget = null;
  el.jumpBar.classList.remove('show');
  clearTimeout(hideJumpBar._t);
  hideJumpBar._t = setTimeout(() => { el.jumpBar.hidden = true; }, 240);
}

function restorePos(k) {
  const pane = PANES[k];
  if (!pane || !pane.entry || !pane.stage) return;
  const r = S_POS.get(pane.entry.p);
  if (r == null || r <= 0.02) return;
  jumpTarget = { k, r };
  el.jumpText.textContent = '上次读到 ' + Math.round(r * 100) + '%';
  el.jumpBar.hidden = false;
  void el.jumpBar.offsetHeight;
  el.jumpBar.classList.add('show');
  clearTimeout(hideJumpBar._t);
  hideJumpBar._t = setTimeout(hideJumpBar, 9000);
}

function fitPane(k, mode) {
  const pane = PANES[k];
  const v = pane.view;
  if (!v) return Promise.resolve();
  const f = v.fitZoom();
  S.fit = f;
  let z = f;
  if (mode === 'page' && S.currentPageH) z = Math.min(f, S.currentPageH);
  return Promise.resolve(v.setZoom(z)).then(() => {
    if (k === ACT) { updateZoomLabel(); pane.stage.scrollTop = 0; }
  });
}

function syncReaderChrome() {
  const f = curEntry();
  if (f) {
    el.rName.textContent = f.n;
    el.rMeta.textContent = `${(TYPE[f.e] || TYPE.pdf).name} · ${fmtSize(f.s)} · ${fmtDate(f.t)}`;
  }
}

async function openEntry(f, paneKey) {
  const k = paneKey || ACT;
  ACT = k;
  assignEntry(k, f);
  S.busy = true;
  syncReaderChrome();
  showReader(true);
  el.loading.hidden = false;
  el.loadingTxt.textContent = '正在打开…';
  const st = PANES[k].stage;
  st.innerHTML = '';
  el.pager.hidden = true;
  if (PANES[k].view) { try { PANES[k].view.destroy(); } catch (e) {} assignView(k, null); }

  const openTask = (async () => {
    if (f.e === 'pdf') await openPdf(f, k);
    else if (f.e === 'md' || f.e === 'markdown') await openMarkdown(f, k);
    else await openOffice(f, k);
  })();
  try {
    await Promise.race([
      openTask,
      new Promise((_, rej) => setTimeout(() => rej(new Error('打开超时（超过 40 秒）')), 40000)),
    ]);
  } catch (e) {
    console.error(e);
    el.loading.hidden = true;
    st.innerHTML = `<div class="errbox"><h3>无法打开这个文件</h3><p>${esc(e && e.message ? e.message : '解析失败')}</p><p class="hint">${f.e === 'doc' || f.e === 'xls' || f.e === 'ppt' ? '这是旧版二进制格式，请先用 Office 另存为 .docx / .xlsx / .pptx 后再打开。' : ''}</p></div>`;
    assignView(k, null);
  } finally {
    S.busy = false;
    el.loading.hidden = true;
  }
  if (k === ACT) { syncReaderChrome(); updatePager(); }
  restorePos(k);
}

async function openPdf(f, k) {
  const st = PANES[k].stage;
  S.kind = 'pdf';
  el.loadingTxt.textContent = '正在载入 PDF…';
  const view = new PdfDocView(st, fileUrl(f.p), (s) => {
    if (s.progress != null) el.loadingTxt.textContent = `正在载入 PDF… ${Math.min(100, Math.round(s.progress * 100))}%`;
    if (s.pages) el.loadingTxt.textContent = '正在排版页面…';
  });
  await view.load();
  assignView(k, view);
  await fitPane(k, 'width');
  if (k === ACT) el.loading.hidden = true;
}

function openLegacy(buf, ext) {
  if (ext === 'doc') {
    const d = parseDoc(buf);
    return {
      kind: 'docx',
      legacy: 'doc',
      html: docHtml(d),
      pageW: 760,
    };
  }
  if (ext === 'xls') {
    const x = parseXls(buf);
    return { kind: 'xlsx', legacy: 'xls', sheets: x.sheets };
  }
  const p = parsePpt(buf);
  return { kind: 'pptx', legacy: 'ppt', pages: p.pages, pageW: p.pageW, pageH: p.pageH };
}

async function openMarkdown(f, k) {
  const st = PANES[k].stage;
  const resp = await fetch(fileUrl(f.p));
  if (!resp.ok) throw new Error('读取文件失败（HTTP ' + resp.status + '）');
  const text = await resp.text();
  el.loadingTxt.textContent = '正在排版 Markdown…';
  const host = document.createElement('div');
  host.className = 'office-host';
  host.innerHTML = renderMarkdown(text);
  st.innerHTML = '';
  st.appendChild(host);
  const sheet = host.querySelector('.md-sheet');
  assignView(k, new OfficeView(st, { kind: 'docx', pageCount: 1 }, host, sheet ? sheet.offsetWidth : 780));
  await fitPane(k, 'width');
  if (k === ACT) el.loading.hidden = true;
}

async function openOffice(f, k) {
  const st = PANES[k].stage;
  const resp = await fetch(fileUrl(f.p));
  if (!resp.ok) throw new Error('读取文件失败（HTTP ' + resp.status + '）');
  if (f.s > 64 * 1024 * 1024) {
    throw new Error('文件较大（' + fmtSize(f.s) + '），文档版式解析需要较多内存，暂不支持打开');
  }
  const buf = await resp.arrayBuffer();
  const head = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
  const isZip = head[0] === 0x50 && head[1] === 0x4b;
  const isOle = head[0] === 0xd0 && head[1] === 0xcf;

  const ext = f.e;
  let doc = null;
  if (isOle) {
    el.loadingTxt.textContent = '正在还原旧版 ' + ext.toUpperCase() + ' 的文字…';
    await new Promise((r) => setTimeout(r, 16));
    doc = openLegacy(buf, ext);
  } else {
    if (!isZip) throw new Error('文件内容无法识别');
    el.loadingTxt.textContent = '正在解析文档…';
    await new Promise((r) => setTimeout(r, 16));
    if (ext === 'docx') doc = renderDocx(buf);
    else if (ext === 'xlsx') doc = renderXlsx(buf);
    else doc = renderPptx(buf);
  }
  S.kind = doc.kind;

  const host = document.createElement('div');
  host.className = 'office-host';

  if (doc.kind === 'docx') {
    host.innerHTML = doc.html;
    st.innerHTML = '';
    st.appendChild(host);
    const sheet = host.querySelector('.doc-sheet');
    const w = sheet ? sheet.offsetWidth : doc.pageW;
    assignView(k, new OfficeView(st, doc, host, w));
  } else if (doc.kind === 'pptx') {
    let html = '';
    for (const pg of doc.pages) html += `<div class="ppt-wrap" style="width:${pg.w}px;height:${pg.h}px">${pg.html}</div>`;
    host.innerHTML = html;
    st.innerHTML = '';
    st.appendChild(host);
    fitPptText(host);
    assignView(k, new OfficeView(st, doc, host, doc.pageW));
  } else {
    const tabs = doc.sheets.map((s, i) => `<button class="xls-tab${i === 0 ? ' is-on' : ''}" data-i="${i}">${esc(s.name)}</button>`).join('');
    host.innerHTML = `<div class="xls-tabs">${tabs}</div><div class="xls-pane">${doc.sheets[0].html}</div>`;
    st.innerHTML = '';
    st.appendChild(host);
    const paneEl = host.querySelector('.xls-pane');
    host.querySelector('.xls-tabs').addEventListener('click', (ev) => {
      const b = ev.target.closest('.xls-tab');
      if (!b) return;
      host.querySelectorAll('.xls-tab').forEach((x) => x.classList.toggle('is-on', x === b));
      const sh = doc.sheets[+b.dataset.i];
      paneEl.innerHTML = sh.html;
      const v = PANES[k].view;
      if (v) { v.baseW = sh.w; v.setZoom(v.fitZoom()); if (k === ACT) updateZoomLabel(); }
    });
    assignView(k, new OfficeView(st, doc, host, doc.sheets[0].w));
  }

  await fitPane(k, 'width');
  if (k === ACT) el.loading.hidden = true;
}

/* ---------------------------------------------------- 书库（合集） */
async function apiCollections() {
  const r = await fetch('/api/collections?t=' + encodeURIComponent(TOKEN), { cache: 'no-store' });
  return r.json();
}

async function loadCollections() {
  try {
    const d = await apiCollections();
    S.collections = (d.collections || []).filter((c) => c.name !== PREFS_COL).map((c) => ({
      id: String(c.id || ''),
      name: String(c.name || '未命名书库'),
      items: Array.isArray(c.items) ? c.items.map(String) : [],
    }));
  } catch (e) {
    S.collections = S.collections || [];
  }
  if (S.collectionId && !S.collections.some((c) => c.id === S.collectionId)) S.collectionId = null;
  renderShelves();
  renderColbar();
}

function liveCount(c) {
  let n = 0;
  for (const p of c.items) if (S.fileMap.has(p)) n++;
  return n;
}

function renderShelves() {
  if (S.collectionId || !S.collections.length) {
    el.shelves.hidden = true;
    el.shelves.innerHTML = '';
    return;
  }
  el.shelves.hidden = false;
  el.shelves.innerHTML = S.collections.map((c) =>
    `<button class="shelf-card" data-id="${esc(c.id)}">
      <svg class="sc-ico" viewBox="0 0 24 24" fill="none"><path d="M4 5.6h5.4L12 7.4h8v10.8a1.3 1.3 0 0 1-1.3 1.3H5.3A1.3 1.3 0 0 1 4 18.2V5.6Z" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M4.4 9.2h15.2" stroke="currentColor" stroke-width="1.6"/></svg>
      <span class="sc-name">${esc(c.name)}</span>
      <span class="sc-count">${liveCount(c)} 份文档</span>
    </button>`).join('');
}

function renderColbar() {
  const c = S.collections.find((x) => x.id === S.collectionId);
  el.colbar.hidden = !c;
  if (c) el.colTitle.textContent = c.name;
}

function openCollection(id) {
  S.collectionId = id;
  S.editMode = false;
  renderShelves();
  renderColbar();
  renderList();
  el.list.scrollTop = 0;
}

function exitCollection() {
  S.collectionId = null;
  renderShelves();
  renderColbar();
  renderList();
  el.list.scrollTop = 0;
}

function showCollectionMenu(c) {
  showSheet('书库：' + c.name, [
    { label: '打开', run: () => openCollection(c.id) },
    { label: '重命名', run: () => renameCollection(c) },
    { label: '解散书库', danger: true, run: () => dissolveCollection(c) },
  ]);
}

async function renameCollection(c) {
  const name = await promptDialog({
    title: '重命名书库',
    text: `当前名称：${c.name}`,
    value: c.name,
    placeholder: '输入新的名字',
    okLabel: '保存',
  });
  if (!name || name === c.name) return;
  try {
    const r = await fetch('/api/collection/rename?t=' + encodeURIComponent(TOKEN) +
      '&id=' + encodeURIComponent(c.id) + '&name=' + encodeURIComponent(name), { method: 'POST' });
    const d = await r.json();
    if (d && d.ok) { c.name = name; renderShelves(); renderColbar(); toast('已重命名为「' + name + '」'); }
    else toast('重命名失败：可能有同名书库');
  } catch (e) { toast('重命名失败'); }
}

async function dissolveCollection(c) {
  const ok = await confirmDialog({
    title: '解散这个书库？',
    text: `「${c.name}」· ${liveCount(c)} 份文档`,
    hint: '只会删除这个分组，书里的文档本身不会被删除。',
    okLabel: '解散',
    danger: true,
  });
  if (!ok) return;
  try {
    const r = await fetch('/api/collection/delete?t=' + encodeURIComponent(TOKEN) +
      '&id=' + encodeURIComponent(c.id), { method: 'POST' });
    const d = await r.json();
    if (d && d.ok) {
      if (S.collectionId === c.id) S.collectionId = null;
      await loadCollections();
      renderList();
      toast('书库已解散，文档保留');
    } else toast('解散失败');
  } catch (e) { toast('解散失败'); }
}

/* ---------------------------------------------------- 多选 */
function toggleSelect(p, node) {
  if (S.selected.has(p)) S.selected.delete(p);
  else S.selected.add(p);
  if (node) node.classList.toggle('on', S.selected.has(p));
  updateEditBar();
}

function updateEditBar() {
  const n = S.selected.size;
  el.ebTitle.textContent = `已选 ${n} 项`;
  el.btnMerge.disabled = n === 0;
  el.btnDelSel.disabled = n === 0;
}

function enterEdit() {
  if (!S.files.length) return;
  S.editing = true;
  S.selected.clear();
  el.shelf.classList.add('editing');
  el.list.classList.add('editing');
  el.topbar.hidden = true;
  el.editbar.hidden = false;
  el.actionbar.hidden = false;
  el.list.querySelectorAll('.item.on').forEach((n) => n.classList.remove('on'));
  updateEditBar();
}

function exitEdit() {
  S.editing = false;
  S.selected.clear();
  el.shelf.classList.remove('editing');
  el.list.classList.remove('editing');
  el.topbar.hidden = false;
  el.editbar.hidden = true;
  el.actionbar.hidden = true;
  el.btnSelect.hidden = !S.files.length;
  el.list.querySelectorAll('.item.on').forEach((n) => n.classList.remove('on'));
}

function mergeSelected() {
  const paths = Array.from(S.selected);
  if (!paths.length) { toast('请先勾选文档'); return; }
  showAddToCollectionSheet(paths);
}

/** 选择「追加到哪个书库」，或新建 */
function showAddToCollectionSheet(paths) {
  if (!paths.length) { toast('请先勾选文档'); return; }
  const items = S.collections.map((c) => ({
    label: `追加到「${c.name}」· ${liveCount(c)} 份`,
    run: () => addFilesToCollection(c.name, paths),
  }));
  items.push({ label: '＋ 新建书库…', run: () => createCollectionFlow(paths) });
  showSheet(`加入书库 · 已选 ${paths.length} 份`, items);
}

async function addFilesToCollection(name, paths) {
  try {
    const r = await fetch('/api/collection/create?t=' + encodeURIComponent(TOKEN), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, items: paths }),
    });
    const d = await r.json();
    if (d && d.ok) {
      if (S.editing) exitEdit();
      await loadCollections();
      renderList();
      toast(`已添加 ${paths.length} 份到「${name}」`, 2600);
    } else {
      toast('添加失败');
    }
  } catch (e) {
    toast('添加失败：' + (e.message || e));
  }
}

async function createCollectionFlow(paths) {
  const name = await promptDialog({
    title: '新建书库',
    text: `已选 ${paths.length} 份文档`,
    hint: '同一个名字的书库会自动合并，不会重复创建',
    placeholder: '给书库起个名字，例如：电子线路讲义',
    okLabel: '创建',
  });
  if (!name) return;
  await addFilesToCollection(name, paths);
}

async function deleteSelected() {
  const paths = Array.from(S.selected);
  if (!paths.length) return;
  const names = paths.slice(0, 5).map((p) => (S.fileMap.get(p) ? S.fileMap.get(p).n : p));
  const ok = await confirmDialog({
    title: `删除 ${paths.length} 个文件？`,
    text: names.join('\n') + (paths.length > 5 ? `\n…等 ${paths.length} 个文件` : ''),
    hint: '将从手机存储中彻底删除，不会进入回收站，删除后无法恢复。',
    okLabel: `删除 ${paths.length} 个`,
    danger: true,
  });
  if (!ok) return;
  let done = 0, fail = 0;
  for (const p of paths) {
    try {
      const r = await fetch('/api/delete?p=' + encodeURIComponent(p) + '&t=' + encodeURIComponent(TOKEN), { method: 'POST' });
      const d = await r.json();
      if (d && d.ok) { done++; S.files = S.files.filter((x) => x.p !== p); }
      else fail++;
    } catch (e) { fail++; }
  }
  exitEdit();
  await loadCollections();
  renderList();
  toast(`已删除 ${done} 个${fail ? `，${fail} 个失败` : ''}`, 2800);
}

/* ---------------------------------------------------- 长按 */
function attachLongPress(container, selector, cb) {
  let timer = 0;
  const arm = (target) => {
    const node = target && target.closest ? target.closest(selector) : null;
    if (!node) return;
    clearTimeout(timer);
    suppressClick = false;
    timer = setTimeout(() => {
      suppressClick = true;
      setTimeout(() => { suppressClick = false; }, 900);
      cb(node);
    }, 480);
  };
  container.addEventListener('touchstart', (e) => { if (e.touches.length === 1) arm(e.target); }, { passive: true });
  container.addEventListener('touchmove', () => clearTimeout(timer), { passive: true });
  container.addEventListener('touchend', () => clearTimeout(timer), { passive: true });
  container.addEventListener('touchcancel', () => clearTimeout(timer), { passive: true });
  container.addEventListener('mousedown', (e) => { if (e.button === 0) arm(e.target); });
  container.addEventListener('mouseup', () => clearTimeout(timer));
  container.addEventListener('contextmenu', (e) => {
    const node = e.target.closest ? e.target.closest(selector) : null;
    if (!node) return;
    e.preventDefault();
    clearTimeout(timer);
    cb(node);
  });
}

/* ------------------------------------------------------------------ 缩放 */
async function fitTo(mode) {
  return fitPane(ACT, mode);
}

function updateZoomLabel() {
  const v = curView();
  if (!v) { el.zVal.textContent = '100%'; return; }
  const pct = Math.round((v.zoom / (S.fit || 1)) * 100);
  el.zVal.textContent = pct + '%';
}

async function zoomBy(factor, anchor) {
  const v = curView();
  if (!v) return;
  const f = v.fitZoom();
  S.fit = f;
  const z0 = v.zoom;
  const z1 = Math.max(f * 0.45, Math.min(f * 10, z0 * factor));
  if (Math.abs(z1 - z0) < 1e-4) return;
  const st = curStage();
  const mx = anchor ? anchor.x : st.clientWidth / 2;
  const my = anchor ? anchor.y : st.clientHeight / 2;
  const cx = (st.scrollLeft + mx) / z0;
  const cy = (st.scrollTop + my) / z0;
  await v.setZoom(z1);
  st.scrollLeft = cx * z1 - mx;
  st.scrollTop = cy * z1 - my;
  updateZoomLabel();
}

/* ------------------------------------------------------------------ 手势 */
function attachGestures(stage, key) {
  const PV = () => PANES[key].view;
  const act = () => setActivePane(key);
  let pinch = null;
  let lastTap = 0, tapX = 0, tapY = 0;

  const dist = (t) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
  const mid = (t) => {
    const r = stage.getBoundingClientRect();
    return { x: (t[0].clientX + t[1].clientX) / 2 - r.left, y: (t[0].clientY + t[1].clientY) / 2 - r.top };
  };

  stage.addEventListener('touchstart', (e) => {
    if (!S.busy) act();
    if (e.touches.length === 2 && PV()) {
      const c = PV().contentEl;
      if (!c) return;
      pinch = {
        d0: dist(e.touches), z0: PV().zoom, mid: mid(e.touches),
        cx: 0, cy: 0, k: 1, el: c,
      };
      const st = stage;
      pinch.cx = (st.scrollLeft + pinch.mid.x) / pinch.z0;
      pinch.cy = (st.scrollTop + pinch.mid.y) / pinch.z0;
      c.style.transition = 'none';
      // CSS zoom 缩放布局时，transform-origin 用的是未缩放的局部坐标
      const ox = PV().layoutZoomed ? pinch.cx : pinch.cx * pinch.z0;
      const oy = PV().layoutZoomed ? pinch.cy : pinch.cy * pinch.z0;
      c.style.transformOrigin = `${ox}px ${oy}px`;
    } else if (e.touches.length === 1) {
      const now = Date.now();
      const t = e.touches[0];
      if (now - lastTap < 300 && Math.abs(t.clientX - tapX) < 26 && Math.abs(t.clientY - tapY) < 26) {
        lastTap = 0;
        onDoubleTap(t);
      } else {
        lastTap = now;
        tapX = t.clientX;
        tapY = t.clientY;
      }
    }
  }, { passive: true });

  stage.addEventListener('touchmove', (e) => {
    if (pinch && e.touches.length >= 2) {
      const k = Math.max(0.2, Math.min(8, dist(e.touches) / pinch.d0));
      pinch.k = k;
      pinch.el.style.transform = `scale(${k})`;
      if (e.cancelable) e.preventDefault();
    }
  }, { passive: false });

  const endPinch = async () => {
    if (!pinch) return;
    const p = pinch;
    pinch = null;
    const v = PV();
    if (!v) { p.el.style.transform = ''; return; }
    const f = v.fitZoom();
    S.fit = f;
    const z1 = Math.max(f * 0.45, Math.min(f * 10, p.z0 * p.k));
    await v.setZoom(z1);
    p.el.style.transform = '';
    p.el.style.transformOrigin = '';
    stage.scrollLeft = p.cx * z1 - p.mid.x;
    stage.scrollTop = p.cy * z1 - p.mid.y;
    requestAnimationFrame(() => {
      stage.scrollLeft = p.cx * z1 - p.mid.x;
      stage.scrollTop = p.cy * z1 - p.mid.y;
    });
    updateZoomLabel();
  };

  stage.addEventListener('touchend', (e) => {
    if (pinch && e.touches.length < 2) endPinch();
  });
  stage.addEventListener('touchcancel', () => { if (pinch) endPinch(); });

  function onDoubleTap(t) {
    const v = PV();
    if (!v) return;
    const f = v.fitZoom();
    S.fit = f;
    const r = stage.getBoundingClientRect();
    const anchor = { x: t.clientX - r.left, y: t.clientY - r.top };
    if (v.zoom > f * 1.25) {
      Promise.resolve(v.setZoom(f)).then(() => { updateZoomLabel(); stage.scrollTop = 0; });
    } else {
      zoomBy(f * 2.4 / v.zoom, anchor);
    }
  }

  // 桌面/外接鼠标：Ctrl + 滚轮缩放
  stage.addEventListener('wheel', (e) => {
    act();
    if (!e.ctrlKey) return;
    e.preventDefault();
    const r = stage.getBoundingClientRect();
    zoomBy(e.deltaY < 0 ? 1.12 : 1 / 1.12, { x: e.clientX - r.left, y: e.clientY - r.top });
  }, { passive: false });
}

/* ------------------------------------------------------------------ 翻页 */
function updatePager() {
  const v = curView();
  if (!v) { el.pager.hidden = true; return; }
  if (v.type === 'xlsx') { el.pager.hidden = true; return; }
  const total = v.total || 1;
  el.pager.hidden = total <= 1;
  const cur = Math.min(total, Math.max(1, v.currentPage()));
  el.pNow.textContent = cur + ' / ' + total;
  el.pPrev.disabled = cur <= 1;
  el.pNext.disabled = cur >= total;
}

let scrollRaf = 0;
function onStageScroll() {
  if (scrollRaf) return;
  scrollRaf = requestAnimationFrame(() => {
    scrollRaf = 0;
    updatePager();
  });
}

/* ------------------------------------------------------------------ 事件绑定 */
function bind() {
  el.list.addEventListener('click', (e) => {
    const it = e.target.closest('.item');
    if (!it) return;
    if (suppressClick) { suppressClick = false; return; }
    if (S.editing) { toggleSelect(it.dataset.p, it); return; }
    const f = S.files.find((x) => x.p === it.dataset.p);
    if (f) openEntry(f);
  });

  // 长按 / 右键
  attachLongPress(el.list, '.item', (node) => {
    if (S.editing) return;
    const f = S.files.find((x) => x.p === node.dataset.p);
    if (f) showFileMenu(f);
  });
  attachLongPress(el.shelves, '.shelf-card', (node) => {
    const c = S.collections.find((x) => x.id === node.dataset.id);
    if (c) showCollectionMenu(c);
  });

  // 书库
  el.shelves.addEventListener('click', (e) => {
    const c = e.target.closest('.shelf-card');
    if (!c) return;
    if (suppressClick) { suppressClick = false; return; }
    openCollection(c.dataset.id);
  });
  el.btnColBack.addEventListener('click', exitCollection);
  el.btnColMenu.addEventListener('click', () => {
    const c = S.collections.find((x) => x.id === S.collectionId);
    if (c) showCollectionMenu(c);
  });

  // 多选
  el.btnAdd.addEventListener('click', showAddSheet);

  el.sheetBody.addEventListener('click', (e) => {
    const row = e.target.closest('[data-addp]');
    if (!row) return;
    const p = row.dataset.addp;
    if (S.manual.has(p)) S.manual.delete(p); else S.manual.add(p);
    savePrefs();
    renderAddList();
    renderList();
  });

  el.btnSelect.addEventListener('click', enterEdit);
  el.btnExitEdit.addEventListener('click', exitEdit);
  el.btnSelAll.addEventListener('click', () => {
    const all = filtered();
    const full = all.length > 0 && all.every((f) => S.selected.has(f.p));
    S.selected.clear();
    if (!full) all.forEach((f) => S.selected.add(f.p));
    el.list.querySelectorAll('.item').forEach((n) => n.classList.toggle('on', S.selected.has(n.dataset.p)));
    updateEditBar();
  });
  el.btnMerge.addEventListener('click', mergeSelected);
  el.btnDelSel.addEventListener('click', deleteSelected);

  el.q.addEventListener('input', () => {
    S.q = el.q.value;
    el.qClear.hidden = !S.q;
    renderList();
  });
  el.qClear.addEventListener('click', () => {
    el.q.value = ''; S.q = ''; el.qClear.hidden = true; renderList();
  });

  el.chips.addEventListener('click', (e) => {
    const c = e.target.closest('.chip');
    if (!c) return;
    el.chips.querySelectorAll('.chip').forEach((x) => x.classList.toggle('is-on', x === c));
    S.filter = c.dataset.f;
    renderList();
    el.list.scrollTop = 0;
  });

  el.btnRefresh.addEventListener('click', async () => {
    el.btnRefresh.classList.add('spinning');
    await apiRescan();
    await loadFiles(true);
    setTimeout(() => el.btnRefresh.classList.remove('spinning'), 500);
    toast('已重新扫描');
  });

  el.btnPerm.addEventListener('click', () => {
    if (window.Android && window.Android.requestAllFiles) window.Android.requestAllFiles();
    else toast('请在系统设置中授予“所有文件访问权限”');
  });

  el.btnBack.addEventListener('click', closeReader);
  el.btnMore.addEventListener('click', showReaderMenu);
  el.sheetMask.addEventListener('click', closeSheet);
  el.sheet.addEventListener('click', (e) => {
    if (e.target.closest('[data-close]')) return closeSheet();
    const b = e.target.closest('[data-idx]');
    if (!b) return;
    const item = sheetState.items[+b.dataset.idx];
    closeSheet();
    if (item && item.run) setTimeout(item.run, 70);
  });
  el.dlgMask.addEventListener('click', () => el.dlgCancel.click());

  el.zIn.addEventListener('click', () => zoomBy(1.25));
  el.zOut.addEventListener('click', () => zoomBy(1 / 1.25));
  el.zVal.addEventListener('click', async () => {
    const v = curView();
    if (!v) return;
    const f = v.fitZoom();
    S.fit = f;
    await v.setZoom(f);
    updateZoomLabel();
    curStage().scrollTop = 0;
  });
  el.btnFit.addEventListener('click', () => fitTo('width'));

  el.pPrev.addEventListener('click', () => { const v = curView(); if (v) { v.gotoPage(v.currentPage() - 1); updatePager(); } });
  el.pNext.addEventListener('click', () => { const v = curView(); if (v) { v.gotoPage(v.currentPage() + 1); updatePager(); } });

  [['a', el.stage], ['b', el.stageB]].forEach(([k, st]) => {
    st.addEventListener('scroll', () => {
      if (!S.busy) setActivePane(k);
      onStageScroll();
      schedulePosSave(k);
    }, { passive: true });
  });

  let rzT = 0;
  window.addEventListener('resize', () => {
    clearTimeout(rzT);
    rzT = setTimeout(() => {
      if (SPLIT.on) { refitBoth(); updatePager(); return; }
      const v = curView();
      if (!v) return;
      const wasFit = Math.abs(v.zoom - S.fit) < S.fit * 0.02;
      if (wasFit) fitTo('width');
      updatePager();
    }, 220);
  });

  document.addEventListener('visibilitychange', () => {
    if (!document.hidden && el.shelf.classList.contains('active')) loadFiles(false);
  });

  el.btnSplit.addEventListener('click', openPicker);
  el.psClose.addEventListener('click', closePicker);
  el.pickerMask.addEventListener('click', closePicker);
  el.psOrient.addEventListener('click', toggleOrientation);
  el.psList.addEventListener('click', (e) => {
    const it = e.target.closest('.ps-item');
    if (!it) return;
    const f = S.files.find((x) => x.p === it.dataset.p);
    if (!f) return;
    closePicker();
    startSplit(f);
  });
  el.btnTheme.addEventListener('click', toggleTheme);
  el.jumpGo.addEventListener('click', () => {
    if (jumpTarget) applyJump(jumpTarget.k, jumpTarget.r);
    hideJumpBar();
  });
  el.jumpClose.addEventListener('click', hideJumpBar);

  attachGestures(el.stage, 'a');
  attachGestures(el.stageB, 'b');
}

/* ------------------------------------------------------------------ 分屏 */
function setActivePane(k) {
  if (!PANES[k]) return;
  ACT = k;
  el.paneA.classList.toggle('active', k === 'a');
  el.paneB.classList.toggle('active', k === 'b');
  if (SPLIT.on) { syncReaderChrome(); updatePager(); updateZoomLabel(); }
}

function renderPicker() {
  const list = S.files.filter((f) => !MANUAL_EXTS.has(f.e) || S.manual.has(f.p));
  if (!list.length) { el.psList.innerHTML = '<div class="ps-empty">还没有可选的文档</div>'; return; }
  el.psList.innerHTML = list.map((f) => {
    const ty = TYPE[f.e] || TYPE.pdf;
    return '<button class="item ps-item" data-p="' + esc(f.p) + '">'
      + '<div class="tile ' + ty.cls + '">' + ty.label + '</div>'
      + '<div class="item-main"><div class="item-name">' + esc(f.n.replace(/\.[^.]+$/, '')) + '</div>'
      + '<div class="item-sub"><span>' + fmtSize(f.s) + '</span><i>·</i><span>' + esc(shortDir(f.d)) + '</span></div></div></button>';
  }).join('');
}

function openPicker() {
  renderPicker();
  el.psOrient.textContent = SPLIT.vertical ? '上下' : '左右';
  el.pickerSide.hidden = false;
  el.pickerMask.hidden = false;
  void el.pickerSide.offsetHeight;
  el.pickerSide.classList.add('show');
  el.pickerMask.classList.add('show');
}

function closePicker() {
  el.pickerSide.classList.remove('show');
  el.pickerMask.classList.remove('show');
  clearTimeout(closePicker._t);
  closePicker._t = setTimeout(() => {
    el.pickerSide.hidden = true;
    el.pickerMask.hidden = true;
  }, 320);
}

function toggleOrientation() {
  SPLIT.vertical = !SPLIT.vertical;
  el.psOrient.textContent = SPLIT.vertical ? '上下' : '左右';
  if (SPLIT.on) {
    el.panes.classList.toggle('vertical', SPLIT.vertical);
    setTimeout(refitBoth, 340);
  }
}

function refitBoth() {
  ['a', 'b'].forEach((k) => {
    const v = PANES[k].view;
    if (v) { try { v.setZoom(v.fitZoom()); } catch (e) {} }
  });
  updateZoomLabel();
}

function startSplit(f) {
  const panes = el.panes;
  if (SPLIT.on) {
    openEntry(f, 'b');
    setTimeout(refitBoth, 380);
    return;
  }
  SPLIT.on = true;
  panes.classList.toggle('vertical', SPLIT.vertical);
  panes.classList.remove('split');
  void panes.offsetHeight;
  panes.classList.add('split');
  setActivePane('b');
  openEntry(f, 'b');
  setTimeout(refitBoth, 380);
}

function exitSplit() {
  if (!SPLIT.on) return;
  recordPos('b'); savePrefs();
  SPLIT.on = false;
  el.panes.classList.add('closing');
  el.panes.classList.remove('split');
  setActivePane('a');
  setTimeout(() => {
    el.panes.classList.remove('vertical', 'closing');
    const v = PANES.b.view;
    if (v) { try { v.destroy(); } catch (e) {} }
    PANES.b.view = null;
    PANES.b.entry = null;
    S.viewB = null;
    S.entryB = null;
    el.stageB.innerHTML = '';
    refitBoth();
    syncReaderChrome();
    updatePager();
  }, 320);
}

/* ---------------------------------------------------- 操作菜单 / 二次确认 */
let suppressClick = false;
let lpTimer = 0;
const sheetState = { items: [] };

function showSheet(title, items) {
  sheetState.items = items;
  el.sheetTitle.textContent = title || '';
  el.sheetTitle.hidden = !title;
  el.sheetBody.classList.remove('add-mode');
  el.sheetBody.innerHTML = items.map((it, i) =>
    `<button class="sheet-item${it.danger ? ' danger' : ''}" data-idx="${i}">${esc(it.label)}</button>`
  ).join('') + '<button class="sheet-item" data-close>取消</button>';
  openSheetShell();
}

function openSheetShell() {
  clearTimeout(sheetTimer);
  el.sheet.hidden = false;
  el.sheetMask.hidden = false;
  // 先让关闭态生效一帧，transition 才有起点
  void el.sheet.offsetHeight;
  el.sheet.classList.add('is-open');
  el.sheetMask.classList.add('is-open');
}

/* 添加文档：Markdown 默认不进列表，在这里挑 */
let addCands = [];
let addQuery = '';

function addRowHtml(f) {
  const on = S.manual.has(f.p);
  return `<button class="sheet-item add-row${on ? ' is-on' : ''}" data-addp="${esc(f.p)}">`
    + `<span class="add-check">${on ? '\u2713' : '+'}</span>`
    + `<span class="add-name">${esc(f.n)}</span>`
    + `<span class="add-dir">${esc(shortDir(f.d))}</span></button>`;
}

function renderAddList() {
  const box = el.sheetBody.querySelector('#addList');
  if (!box) return;
  const top = box.scrollTop;
  const q = addQuery.trim().toLowerCase();
  const list = addCands.filter((f) => !q || f.n.toLowerCase().includes(q) || (f.d || '').toLowerCase().includes(q));
  box.innerHTML = list.length
    ? list.map(addRowHtml).join('')
    : '<div class="add-empty">没有匹配的文件</div>';
  box.scrollTop = top;
}

function showAddSheet() {
  addCands = S.files.filter((f) => MANUAL_EXTS.has(f.e)).sort((a, b) => b.t - a.t);
  if (!addCands.length) { toast('没找到可添加的 Markdown 文件'); return; }
  addQuery = '';
  el.sheetTitle.textContent = '添加文档 · 共 ' + addCands.length + ' 个';
  el.sheetTitle.hidden = false;
  el.sheetBody.classList.add('add-mode');
  el.sheetBody.innerHTML = '<div class="add-search"><input id="addQ" type="search" '
    + 'placeholder="搜索文件名或目录" autocomplete="off" spellcheck="false"></div>'
    + '<div class="add-list" id="addList"></div>'
    + '<div class="add-foot"><button class="sheet-item" data-close>完成</button></div>';
  renderAddList();
  el.sheetBody.querySelector('#addQ').addEventListener('input', (e) => {
    addQuery = e.target.value;
    renderAddList();
  });
  openSheetShell();
}

let sheetTimer = 0;
function closeSheet() {
  if (el.sheet.hidden) return;
  el.sheet.classList.remove('is-open');
  el.sheetMask.classList.remove('is-open');
  clearTimeout(sheetTimer);
  const reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
  sheetTimer = setTimeout(() => {
    el.sheet.hidden = true;
    el.sheetMask.hidden = true;
  }, reduce ? 180 : 260);
}

function showFileMenu(f) {
  const items = [
    { label: '打开', run: () => openEntry(f) },
    { label: '文件信息', run: () => toast(`${f.n}\n${fmtSize(f.s)} · ${fmtDate(f.t)}\n${f.d}`, 5200) },
    {
      label: '用其他应用打开',
      run: () => {
        if (window.Android && window.Android.openPath) window.Android.openPath(f.p);
        else toast('仅 App 内支持');
      },
    },
    { label: '加入书库…', run: () => showAddToCollectionSheet([f.p]) },
    { label: '删除文件', danger: true, run: () => deleteFile(f) },
  ];
  if (MANUAL_EXTS.has(f.e)) {
    items.splice(items.length - 1, 0, {
      label: '从添加列表移除',
      run: () => { S.manual.delete(f.p); savePrefs(); renderList(); toast('已移出添加列表'); },
    });
  }
  showSheet(f.n, items);
}

function showReaderMenu() {
  const f = curEntry();
  if (!f) return;
  const menu = [
    { label: '文件信息', run: () => toast(`${f.n}\n${fmtSize(f.s)} · ${fmtDate(f.t)}\n${f.d}`, 5200) },
    {
      label: '用其他应用打开',
      run: () => {
        if (window.Android && window.Android.openPath) window.Android.openPath(f.p);
        else toast('仅 App 内支持');
      },
    },
    { label: '重新加载', run: () => openEntry(f) },
    { label: '加入书库…', run: () => showAddToCollectionSheet([f.p]) },
    { label: '删除文件', danger: true, run: () => deleteFile(f) },
  ];
  if (SPLIT.on) menu.push({ label: '退出分屏', run: exitSplit });
  showSheet(f.n, menu);
}

/** 对话框内核：withInput=false 返回 boolean，true 返回输入内容或 null */
function dialogCore(opts) {
  return new Promise((resolve) => {
    el.dlgTitle.textContent = opts.title || '确认操作';
    el.dlgText.innerHTML = esc(opts.text || '').replace(/\n/g, '<br>') +
      (opts.hint ? `<span class="dlg-hint">${esc(opts.hint)}</span>` : '');
    el.dlgInput.hidden = !opts.withInput;
    if (opts.withInput) {
      el.dlgInput.value = opts.value || '';
      el.dlgInput.placeholder = opts.placeholder || '';
      el.dlgInput.maxLength = opts.maxLength || 40;
    }
    el.dlgOk.textContent = opts.okLabel || '确定';
    el.dlgOk.classList.toggle('danger', !!opts.danger);
    el.dlg.hidden = false;
    el.dlgMask.hidden = false;
    if (opts.withInput) {
      setTimeout(() => { try { el.dlgInput.focus(); el.dlgInput.select(); } catch (e) {} }, 130);
    }

    const finish = (v) => {
      el.dlg.hidden = true;
      el.dlgMask.hidden = true;
      el.dlgOk.removeEventListener('click', onOk);
      el.dlgCancel.removeEventListener('click', onCancel);
      el.dlgInput.removeEventListener('keydown', onKey);
      resolve(v);
    };
    const onOk = () => {
      if (!opts.withInput) return finish(true);
      const v = el.dlgInput.value.trim();
      if (!v) { toast('名字不能为空'); try { el.dlgInput.focus(); } catch (e) {} return; }
      finish(v);
    };
    const onCancel = () => finish(opts.withInput ? null : false);
    const onKey = (e) => { if (e.key === 'Enter') { e.preventDefault(); onOk(); } };
    el.dlgOk.addEventListener('click', onOk);
    el.dlgCancel.addEventListener('click', onCancel);
    el.dlgInput.addEventListener('keydown', onKey);
  });
}

function confirmDialog(opts) { return dialogCore({ ...opts, withInput: false }); }
function promptDialog(opts) { return dialogCore({ ...opts, withInput: true }); }

function deleteFailText(code) {
  if (code === 1) return '该文件不在文档索引中，已跳过';
  if (code === 2) return '没有写入权限，删除失败';
  if (code === 3) return '系统拒绝删除该文件';
  return '删除失败';
}

async function deleteFile(f) {
  const ok = await confirmDialog({
    title: '删除这个文件？',
    text: `${f.n}\n${shortDir(f.d)} · ${fmtSize(f.s)}`,
    hint: '将从手机存储中彻底删除，不会进入回收站，删除后无法恢复。',
    okLabel: '删除',
    danger: true,
  });
  if (!ok) return;
  try {
    const r = await fetch('/api/delete?p=' + encodeURIComponent(f.p) + '&t=' + encodeURIComponent(TOKEN), { method: 'POST' });
    const d = await r.json();
    if (d && d.ok) {
      S.files = S.files.filter((x) => x.p !== f.p);
      toast('已删除 ' + f.n, 2600);
      const isOpen = (PANES.a.entry && PANES.a.entry.p === f.p) || (PANES.b.entry && PANES.b.entry.p === f.p);
      if (isOpen) {
        closeReader();
      } else {
        renderList();
        renderShelves();
      }
    } else {
      toast(deleteFailText(d && d.code), 3000);
    }
  } catch (e) {
    toast('删除失败：' + (e.message || e), 3000);
  }
}

function closeReader() {
  recordPos('a'); recordPos('b'); savePrefs(); hideJumpBar();
  ['a', 'b'].forEach((k) => {
    const p = PANES[k];
    if (p.view) { try { p.view.destroy(); } catch (e) {} }
    p.view = null;
    p.entry = null;
    if (p.stage) p.stage.innerHTML = '';
  });
  S.view = null; S.viewB = null;
  S.entry = null; S.entryB = null;
  SPLIT.on = false;
  el.panes.classList.remove('split', 'vertical');
  el.pickerSide.classList.remove('show');
  el.pickerMask.classList.remove('show');
  el.pickerSide.hidden = true;
  el.pickerMask.hidden = true;
  showReader(false);
  loadFiles(false);
}

/* ------------------------------------------------------------------ 原生回调 */
window.__appBack = function () {
  if (!el.dlg.hidden) { el.dlgCancel.click(); return true; }
  if (!el.sheet.hidden) { closeSheet(); return true; }
  if (!el.pickerSide.hidden) { closePicker(); return true; }
  if (SPLIT.on) { exitSplit(); return true; }
  if (el.reader.classList.contains('active')) { closeReader(); return true; }
  if (S.editing) { exitEdit(); return true; }
  if (S.collectionId) { exitCollection(); return true; }
  return false;
};
window.__onNativeResume = function () {
  const perm = window.Android && window.Android.hasAllFiles ? window.Android.hasAllFiles() : true;
  el.permBar.hidden = !!perm;
  if (el.shelf.classList.contains('active')) loadFiles(false, 4);
};

window.__handleBack = window.__appBack;

/* ------------------------------------------------------------------ 启动 */
async function boot() {
  PANES.a.stage = el.stage;
  PANES.b.stage = el.stageB;
  initTheme();
  await loadPrefs();
  bind();
  const perm = window.Android && window.Android.hasAllFiles ? window.Android.hasAllFiles() : true;
  el.permBar.hidden = !!perm;
  if (!perm) setTimeout(() => toast('请授予“所有文件访问权限”，以便自动收录文档', 3600), 600);
  await loadCollections();
  await loadFiles(true, 4);
  if (window.Android && window.Android.deviceInfo) {
    try { console.log('设备：' + window.Android.deviceInfo()); } catch (e) {}
  }
}

/* 诊断面板：?debug=1 时显示运行指标 */
function initDebug() {
  if (!new URLSearchParams(location.search).has('debug')) return;
  const d = document.createElement('div');
  d.id = 'diag';
  d.style.cssText = 'position:fixed;left:0;top:0;z-index:9999;background:rgba(0,0,0,.82);color:#7df;font:10.5px/1.45 monospace;padding:5px 8px;max-width:100%;white-space:pre-wrap;pointer-events:none';
  document.body.appendChild(d);
  setInterval(() => {
    const v = S.view;
    const st = el.stage;
    const lines = [];
    lines.push('view=' + (S.kind || '-') + ' zoom=' + (v ? v.zoom.toFixed(3) : '-') + ' fit=' + S.fit.toFixed(3));
    lines.push('stage ' + st.clientWidth + 'x' + st.clientHeight + ' scroll ' + st.scrollWidth + 'x' + st.scrollHeight);
    if (S.kind === 'pdf' && v) {
      const pages = st.querySelectorAll('.pdf-page');
      const painted = st.querySelectorAll('.pdf-page.painted').length;
      const c0 = st.querySelector('.pdf-canvas');
      lines.push('pages=' + pages.length + ' painted=' + painted + ' total=' + v.total);
      if (c0) {
        lines.push('canvas buf=' + c0.width + 'x' + c0.height + ' css=' + c0.style.width + 'x' + c0.style.height);
        const r = c0.getBoundingClientRect();
        lines.push('canvas rect=' + Math.round(r.width) + 'x' + Math.round(r.height));
      }
      lines.push('tried=' + (v.viewer.renderTried || 0) + ' err=' + (v.viewer.lastError || '-'));
      lines.push('range=' + JSON.stringify(v.viewer.lastRange || null));
    } else if (v && v.host) {
      const host = v.host;
      lines.push('host ' + Math.round(host.getBoundingClientRect().width) + 'x' + Math.round(host.getBoundingClientRect().height) + ' base=' + Math.round(v.baseW));
      const sheet = host.querySelector('.doc-sheet, .ppt-wrap, .xls-host');
      if (sheet) lines.push('sheet ' + Math.round(sheet.getBoundingClientRect().width));
    }
    lines.push('dpr=' + window.devicePixelRatio + ' vw=' + window.innerWidth + ' vis=' + document.visibilityState);
    d.textContent = lines.join('\n');
  }, 700);
}

window.addEventListener('error', (e) => {
  try { toast('脚本错误：' + (e.message || e.type), 6000); } catch (_) {}
});
window.addEventListener('unhandledrejection', (e) => {
  const m = e.reason && (e.reason.message || e.reason);
  try { toast('错误：' + m, 6000); } catch (_) {}
});

boot();
initDebug();
