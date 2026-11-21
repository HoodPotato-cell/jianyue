/* PDF 查看器：按缩放级别重新栅格化，任意倍数都保持原文件级清晰度 */
const pdfjsLib = globalThis.pdfjsLib;
if (pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
}

const MAX_TILE_PX = 12 * 1024 * 1024; // 单次渲染像素上限，避免爆内存

export class PdfViewer {
  constructor(root, onState) {
    this.root = root;
    this.onState = onState || (() => {});
    this.pdf = null;
    this.zoom = 1;
    this.pages = [];
    this.dpr = Math.min(window.devicePixelRatio || 1, 3);
    this.destroyed = false;
    this._raf = 0;
    this._onScroll = () => this.schedule();
    this.root.addEventListener('scroll', this._onScroll, { passive: true });
    window.addEventListener('resize', this._onScroll);
    this.observer = null;
  }

  async load(url, password) {
    const params = {
      url,
      cMapUrl: 'cmaps/',
      cMapPacked: true,
      standardFontDataUrl: 'standard_fonts/',
      disableAutoFetch: false,
      enableXfa: true,
    };
    const task = pdfjsLib.getDocument(params);
    task.onProgress = (p) => {
      if (p && p.total) this.onState({ progress: p.loaded / p.total });
    };
    this.pdf = await task.promise;
    if (this.destroyed) return;
    this.onState({ pages: this.pdf.numPages, loaded: true });
    await this.buildPages();
  }

  async buildPages() {
    const frag = document.createDocumentFragment();
    this.pages = [];
    // 先取第一页尺寸，其余按需
    for (let i = 1; i <= this.pdf.numPages; i++) {
      const el = document.createElement('div');
      el.className = 'pdf-page';
      el.dataset.n = i;
      const canvas = document.createElement('canvas');
      canvas.className = 'pdf-canvas';
      el.appendChild(canvas);
      const ph = document.createElement('div');
      ph.className = 'pdf-ph';
      ph.textContent = i;
      el.appendChild(ph);
      frag.appendChild(el);
      this.pages.push({ el, canvas, page: null, size: null, rendered: '', task: null, painting: false });
    }
    this.root.innerHTML = '';
    const wrap = document.createElement('div');
    wrap.className = 'pdf-pages';
    wrap.appendChild(frag);
    this.root.appendChild(wrap);
    this.wrap = wrap;

    // 先精确测量前几页，其余按同样尺寸估算，保证首屏立刻可用
    const probe = Math.min(this.pdf.numPages, 8);
    await Promise.all(Array.from({ length: probe }, (_, i) => this.measure(i)));
    const ref = this.pages.find((p) => p.size) || null;
    if (ref) {
      for (const p of this.pages) {
        if (!p.size) { p.size = { w: ref.size.w, h: ref.size.h }; p.estimated = true; }
      }
    }
    this.applyZoom(this.zoom, true);
  }

  async measure(i) {
    const p = this.pages[i];
    if (!p) return null;
    if (p.size && !p.estimated) return p.size;
    const page = p.page || await this.pdf.getPage(i + 1);
    p.page = page;
    const vp = page.getViewport({ scale: 1 });
    const changed = !p.size || Math.abs(p.size.w - vp.width) > 0.5 || Math.abs(p.size.h - vp.height) > 0.5;
    p.size = { w: vp.width, h: vp.height };
    p.estimated = false;
    if (changed) {
      p.el.style.width = (vp.width * this.zoom).toFixed(2) + 'px';
      p.el.style.height = (vp.height * this.zoom).toFixed(2) + 'px';
      p.rendered = '';
    }
    return p.size;
  }

  setZoom(z, opts = {}) {
    this.zoom = Math.max(0.2, Math.min(z, 12));
    this.applyZoom(this.zoom, false, opts);
  }

  async applyZoom(z, initial, opts = {}) {
    this.zoom = z;
    this.wrap.style.setProperty('--zw', z);
    let maxW = 0;
    for (let i = 0; i < this.pages.length; i++) {
      const p = this.pages[i];
      if (!p.size) continue;
      p.el.style.width = (p.size.w * z).toFixed(2) + 'px';
      p.el.style.height = (p.size.h * z).toFixed(2) + 'px';
      maxW = Math.max(maxW, p.size.w * z);
      p.rendered = '';   // 标记需重绘
    }
    this.wrap.style.width = Math.ceil(maxW) + 'px';
    this.wrap.style.margin = '0 auto';
    this.schedule(true);
    if (!initial) this.onState({ zoom: z });
  }

  /** 适应宽度 */
  fitWidth(extra = 24) {
    const avail = this.root.clientWidth - extra;
    const first = this.pages.find((p) => p.size);
    if (!first) return this.zoom;
    return avail / first.size.w;
  }

  schedule(force) {
    if (this._pending) return;
    this._pending = true;
    const run = () => {
      if (!this._pending) return;
      this._pending = false;
      clearTimeout(this._fallback);
      this.paintVisible(force);
    };
    // rAF 在后台标签页会被暂停，加一个定时兜底
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(run);
    this._fallback = setTimeout(run, 120);
  }

  visibleRange() {
    const top = this.root.scrollTop;
    const bottom = top + this.root.clientHeight;
    const head = this.wrap.getBoundingClientRect().top;
    let first = 0, last = this.pages.length - 1;
    for (let i = 0; i < this.pages.length; i++) {
      const p = this.pages[i];
      const y0 = p.el.offsetTop + head;
      const y1 = y0 + p.el.offsetHeight;
      if (y1 > top - 200) { first = i; break; }
    }
    for (let i = first; i < this.pages.length; i++) {
      const p = this.pages[i];
      const y0 = p.el.offsetTop + head;
      if (y0 > bottom + 200) { last = Math.max(first, i - 1); break; }
      last = i;
    }
    return [first, last];
  }

  async paintVisible(force) {
    if (!this.pages.length || this.destroyed) return;
    const [first, last] = this.visibleRange();
    // 取消视野外的渲染任务
    for (let i = 0; i < this.pages.length; i++) {
      if (i < first - 1 || i > last + 1) {
        const p = this.pages[i];
        if (p.task) { try { p.task.cancel(); } catch (e) {} p.task = null; p.rendered = ''; }
      }
    }
    for (let i = Math.max(0, first - 1); i <= Math.min(this.pages.length - 1, last + 1); i++) {
      this.renderPage(i, force).catch((e) => {
        this.lastError = (e && e.message ? e.message : String(e));
        console.error('renderPage', i, e);
      });
    }
    this.lastRange = [first + 1, last + 1];
    this.onState({ range: [first + 1, last + 1] });
  }

  async renderPage(i, force) {
    this.renderTried = (this.renderTried || 0) + 1;
    const p = this.pages[i];
    if (!p || p.painting) { p._again = true; return; }
    if (!p.size || p.estimated) await this.measure(i);
    if (!p.size) return;

    const z = this.zoom;
    const cssW = p.size.w * z;
    const cssH = p.size.h * z;
    const dpr = this.dpr;

    // 计算需要渲染的区域（CSS 像素，相对页面左上角）
    let tileW = cssW, tileH = cssH;
    const needPx = cssW * dpr * cssH * dpr;
    if (needPx > MAX_TILE_PX) {
      // 只渲染视口附近的切片
      const pageTop = p.el.offsetTop + this.wrap.getBoundingClientRect().top;
      const viewTop = this.root.scrollTop;
      const viewH = this.root.clientHeight;
      const cy = Math.max(viewH / 2, Math.min(viewTop + viewH / 2 - pageTop, cssH - viewH / 2));
      const scaleW = Math.sqrt(MAX_TILE_PX * (cssW / cssH)) / dpr;
      tileW = Math.max(120, Math.min(cssW, scaleW));
      tileH = Math.max(120, Math.min(cssH, (MAX_TILE_PX / (dpr * dpr)) / tileW));
      p.tile = { w: tileW, h: tileH };
      p.tileY = Math.max(0, Math.min(cy - tileH / 2, cssH - tileH));
      p.tileX = Math.max(0, Math.min(cssW / 2 - tileW / 2, cssW - tileW));
    } else {
      p.tile = { w: cssW, h: cssH };
      p.tileX = 0;
      p.tileY = 0;
    }

    const key = `${z.toFixed(4)}|${p.tileX.toFixed(0)}|${p.tileY.toFixed(0)}|${dpr}`;
    if (!force && p.rendered === key) return;
    if (p.task) { try { p.task.cancel(); } catch (e) {} p.task = null; }

    p.painting = true;
    p.rendered = key;
    try {
      const page = p.page || await this.pdf.getPage(i + 1);
      p.page = page;
      const vp = page.getViewport({ scale: z * dpr });
      const canvas = p.canvas;
      const tw = Math.max(1, Math.floor(p.tile.w * dpr));
      const th = Math.max(1, Math.floor(p.tile.h * dpr));
      canvas.width = tw;
      canvas.height = th;
      canvas.style.width = p.tile.w.toFixed(2) + 'px';
      canvas.style.height = p.tile.h.toFixed(2) + 'px';
      canvas.style.left = p.tileX.toFixed(2) + 'px';
      canvas.style.top = p.tileY.toFixed(2) + 'px';
      const ctx = canvas.getContext('2d', { alpha: false });
      ctx.save();
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, tw, th);
      ctx.translate(-p.tileX * dpr, -p.tileY * dpr);
      const task = page.render({ canvasContext: ctx, viewport: vp, intent: 'display' });
      p.task = task;
      await task.promise;
      ctx.restore();
      p.el.classList.add('painted');
    } catch (e) {
      if (e && e.name !== 'RenderingCancelledException') p.rendered = '';
    } finally {
      p.painting = false;
      p.task = null;
      if (p._again) { p._again = false; this.schedule(); }
    }
  }

  destroy() {
    this.destroyed = true;
    this.root.removeEventListener('scroll', this._onScroll);
    window.removeEventListener('resize', this._onScroll);
    if (this.observer) this.observer.disconnect();
    for (const p of this.pages) {
      if (p.task) { try { p.task.cancel(); } catch (e) {} }
      if (p.page) { try { p.page.cleanup(); } catch (e) {} }
    }
    if (this.pdf) { try { this.pdf.destroy(); } catch (e) {} }
    this.pages = [];
    this.root.innerHTML = '';
  }

  /** 跳到第 n 页（1 基） */
  gotoPage(n) {
    const p = this.pages[n - 1];
    if (!p) return;
    this.root.scrollTop = p.el.offsetTop;
    this.schedule(true);
  }

  currentPage() {
    const top = this.root.scrollTop + 8;
    for (let i = 0; i < this.pages.length; i++) {
      const p = this.pages[i];
      if (p.el.offsetTop + p.el.offsetHeight > top) return i + 1;
    }
    return this.pages.length;
  }
}
