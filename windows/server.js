'use strict';
/* 简阅 Windows · 本地服务：扫描、文件流、书库、删除 */

const http = require('http');
const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const os = require('os');
const crypto = require('crypto');

const EXT = new Set(['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'md', 'markdown']);
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png', '.jpg': 'image/jpeg', '.svg': 'image/svg+xml',
  '.ttf': 'font/ttf', '.pfb': 'application/octet-stream', '.bcmap': 'application/octet-stream',
  '.pdf': 'application/pdf', '.md': 'text/markdown; charset=utf-8', '.markdown': 'text/markdown; charset=utf-8', '.woff2': 'font/woff2', '.ico': 'image/x-icon',
};

const SKIP_DIRS = new Set([
  'node_modules', '$recycle.bin', 'system volume information', 'appdata', 'windows',
  'program files', 'program files (x86)', 'programdata', '.git', '.cache', '.gradle',
  '.android', 'temp', 'tmp',
]);

const state = {
  roots: [],
  files: [],
  scannedAt: 0,
  scanning: false,
  scanningStartedAt: 0,
  token: '',
  assetsDir: '',
  storeFile: '',
  collections: [],
};

/* ---------------------------------------------------------------- 书库存储 */
async function loadStore() {
  try {
    const raw = await fsp.readFile(state.storeFile, 'utf8');
    const data = JSON.parse(raw);
    state.collections = Array.isArray(data.collections) ? data.collections : [];
    if (Array.isArray(data.roots) && data.roots.length) state.roots = data.roots.map((x) => path.resolve(x));
  } catch (e) {
    state.collections = [];
  }
}
async function saveStore() {
  try {
    await fsp.mkdir(path.dirname(state.storeFile), { recursive: true });
    await fsp.writeFile(state.storeFile, JSON.stringify({ collections: state.collections, roots: state.roots }, null, 2), 'utf8');
  } catch (e) {
    console.error('书库保存失败', e.message);
  }
}

/* ---------------------------------------------------------------- 扫描 */
function isHidden(name) {
  return name.startsWith('.') || name.startsWith('~$');
}

async function scanRoots() {
  if (state.scanning) return;
  state.scanning = true;
  state.scanningStartedAt = Date.now();
  const out = [];
  const seen = new Set();
  const started = Date.now();
  const LIMIT = 20000;
  const MAX_MS = 25000;

  async function walk(dir, depth) {
    if (depth > 10 || out.length >= LIMIT || Date.now() - started > MAX_MS) return;
    let entries;
    try {
      entries = await fsp.readdir(dir, { withFileTypes: true });
    } catch (e) {
      return;
    }
    for (const ent of entries) {
      if (out.length >= LIMIT || Date.now() - started > MAX_MS) return;
      const name = ent.name;
      if (isHidden(name)) continue;
      const full = path.join(dir, name);
      if (ent.isDirectory()) {
        const low = name.toLowerCase();
        if (SKIP_DIRS.has(low)) continue;
        if (/^[a-z]:$/i.test(name)) continue;
        await walk(full, depth + 1);
        continue;
      }
      if (!ent.isFile()) continue;
      const ext = path.extname(name).slice(1).toLowerCase();
      if (!EXT.has(ext)) continue;
      let st;
      try {
        st = await fsp.stat(full);
      } catch (e) {
        continue;
      }
      if (seen.has(full)) continue;
      seen.add(full);
      out.push({
        p: full,
        n: name,
        d: dir.split(path.sep).join('/'),
        s: st.size,
        t: Math.round(st.mtimeMs),
        e: ext,
      });
    }
  }

  for (const root of state.roots) {
    try {
      const st = await fsp.stat(root);
      if (st.isDirectory()) await walk(root, 0);
    } catch (e) { /* 目录不存在就跳过 */ }
  }

  out.sort((a, b) => b.t - a.t);
  state.files = out.slice(0, 8000);
  state.scannedAt = Date.now();
  state.scanning = false;
}

function rescan() {
  if (state.scanning) return;
  scanRoots().catch((e) => {
    console.error('扫描失败', e);
    state.scanning = false;
  });
}

function defaultRoots() {
  const home = os.homedir();
  const list = [
    path.join(home, 'Desktop'),
    path.join(home, 'Documents'),
    path.join(home, 'Downloads'),
  ];
  return list.filter((p) => {
    try { return fs.statSync(p).isDirectory(); } catch (e) { return false; }
  });
}

/* ---------------------------------------------------------------- HTTP */
function send(res, code, body, headers) {
  const buf = Buffer.isBuffer(body) ? body : Buffer.from(String(body));
  res.writeHead(code, Object.assign({ 'Content-Length': buf.length }, headers || {}));
  res.end(buf);
}
function json(res, obj, code) {
  send(res, code || 200, JSON.stringify(obj), { 'Content-Type': 'application/json; charset=utf-8' });
}

function mimeOf(file) {
  return MIME[path.extname(file).toLowerCase()] || 'application/octet-stream';
}

async function serveStatic(req, res, pathname) {
  let rel = decodeURIComponent(pathname);
  if (rel === '/' || rel === '') rel = '/index.html';
  const file = path.join(state.assetsDir, path.normalize(rel).replace(/^(\.\.[/\\])+/, ''));
  if (!file.startsWith(state.assetsDir)) return send(res, 403, 'forbidden');
  let st;
  try { st = await fsp.stat(file); } catch (e) { return send(res, 404, 'not found'); }
  if (!st.isFile()) return send(res, 404, 'not found');
  res.writeHead(200, {
    'Content-Type': mimeOf(file),
    'Content-Length': st.size,
    'Cache-Control': 'no-cache',
  });
  fs.createReadStream(file).pipe(res);
}

async function serveFile(req, res, url) {
  const p = url.searchParams.get('p') || '';
  if (!p) return send(res, 400, 'missing p');
  const abs = path.resolve(p);
  const ext = path.extname(abs).slice(1).toLowerCase();
  if (!EXT.has(ext)) return send(res, 403, 'unsupported');
  let st;
  try { st = await fsp.stat(abs); } catch (e) { return send(res, 404, 'not found'); }
  if (!st.isFile()) return send(res, 404, 'not found');

  const range = req.headers.range;
  const type = mimeOf(abs);
  if (range) {
    const m = /^bytes=(\d*)-(\d*)$/.exec(range.trim());
    if (m) {
      let start = m[1] === '' ? null : parseInt(m[1], 10);
      let end = m[2] === '' ? null : parseInt(m[2], 10);
      if (start === null) { start = Math.max(0, st.size - (end || 0)); end = st.size - 1; }
      if (end === null || end >= st.size) end = st.size - 1;
      if (start > end || start >= st.size) {
        return send(res, 416, '', { 'Content-Range': `bytes */${st.size}` });
      }
      res.writeHead(206, {
        'Content-Type': type,
        'Content-Length': end - start + 1,
        'Content-Range': `bytes ${start}-${end}/${st.size}`,
        'Accept-Ranges': 'bytes',
        'Cache-Control': 'no-store',
      });
      fs.createReadStream(abs, { start, end }).pipe(res);
      return;
    }
  }
  res.writeHead(200, {
    'Content-Type': type,
    'Content-Length': st.size,
    'Accept-Ranges': 'bytes',
    'Cache-Control': 'no-store',
  });
  fs.createReadStream(abs).pipe(res);
}

function readBody(req) {
  return new Promise((resolve) => {
    let data = '';
    req.on('data', (c) => { data += c; if (data.length > 2e6) req.destroy(); });
    req.on('end', () => {
      try { resolve(data ? JSON.parse(data) : {}); } catch (e) { resolve({}); }
    });
  });
}

async function handleApi(req, res, url) {
  const route = url.pathname;
  const method = req.method || 'GET';
  if (url.searchParams.get('t') !== state.token) return json(res, { ok: false, error: 'bad token' }, 403);

  if (route === '/api/files') {
    return json(res, { files: state.files, scanning: state.scanning, scannedAt: state.scannedAt });
  }
  if (route === '/api/rescan') {
    rescan();
    return json(res, { ok: true, scanning: true });
  }
  if (route === '/api/file') {
    return serveFile(req, res, url);
  }
  if (route === '/api/collections') {
    return json(res, { collections: state.collections });
  }
  if (route === '/api/collection/create') {
    const body = await readBody(req);
    const name = String(body.name || '').trim() || '未命名书库';
    const items = Array.isArray(body.items) ? body.items.map(String) : [];
    let col = state.collections.find((c) => c.name === name);
    if (!col) {
      col = { id: crypto.randomUUID(), name, items: [] };
      state.collections.push(col);
    }
    for (const p of items) if (!col.items.includes(p)) col.items.push(p);
    await saveStore();
    return json(res, { ok: true, collection: col });
  }
  if (route === '/api/collection/rename') {
    const id = url.searchParams.get('id');
    const name = String(url.searchParams.get('name') || '').trim();
    const col = state.collections.find((c) => c.id === id);
    if (!col || !name) return json(res, { ok: false });
    if (state.collections.some((c) => c !== col && c.name === name)) return json(res, { ok: false });
    col.name = name;
    await saveStore();
    return json(res, { ok: true });
  }
  if (route === '/api/collection/delete') {
    const id = url.searchParams.get('id');
    const before = state.collections.length;
    state.collections = state.collections.filter((c) => c.id !== id);
    if (state.collections.length !== before) await saveStore();
    return json(res, { ok: state.collections.length !== before });
  }
  if (route === '/api/delete' && method === 'POST') {
    const p = url.searchParams.get('p') || '';
    if (!p) return json(res, { ok: false });
    try {
      await fsp.unlink(path.resolve(p));
      state.files = state.files.filter((f) => f.p !== p);
      for (const c of state.collections) c.items = c.items.filter((x) => x !== p);
      await saveStore();
      return json(res, { ok: true });
    } catch (e) {
      return json(res, { ok: false, error: e.message });
    }
  }
  return json(res, { ok: false, error: 'unknown api' }, 404);
}

function createServer() {
  return http.createServer((req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    if (url.pathname.startsWith('/api/')) {
      handleApi(req, res, url).catch((e) => {
        console.error(e);
        try { json(res, { ok: false, error: e.message }, 500); } catch (_) {}
      });
      return;
    }
    serveStatic(req, res, url.pathname).catch((e) => {
      try { send(res, 500, 'error'); } catch (_) {}
    });
  });
}

async function start(opts) {
  state.assetsDir = path.resolve(opts.assetsDir);
  state.storeFile = path.resolve(opts.storeFile);
  state.token = opts.token || crypto.randomBytes(8).toString('hex');
  state.roots = (opts.roots && opts.roots.length ? opts.roots : defaultRoots()).map((p) => path.resolve(p));

  await loadStore();

  const server = createServer();
  await new Promise((resolve) => server.listen(opts.port || 0, '127.0.0.1', resolve));
  const port = server.address().port;

  rescan();

  return {
    port,
    token: state.token,
    url: `http://127.0.0.1:${port}/index.html?t=${encodeURIComponent(state.token)}`,
    roots: () => state.roots.slice(),
    addRoot(dir) {
      const abs = path.resolve(dir);
      if (!state.roots.includes(abs)) {
        state.roots.push(abs);
        saveStore();
        rescan();
      }
      return state.roots.slice();
    },
    rescan,
    close: () => new Promise((r) => server.close(r)),
  };
}

module.exports = { start };

/* CLI：node server.js --root D:\Docs --port 8792 --token dev */
if (require.main === module) {
  const args = process.argv.slice(2);
  const get = (k) => {
    const i = args.indexOf('--' + k);
    return i >= 0 ? args[i + 1] : undefined;
  };
  const roots = [];
  for (let i = 0; i < args.length; i++) if (args[i] === '--root') roots.push(args[i + 1]);
  start({
    assetsDir: path.join(__dirname, 'assets', 'web'),
    storeFile: path.join(os.homedir(), '.jianyue', 'store.json'),
    roots: roots.length ? roots : defaultRoots(),
    port: parseInt(get('port') || '8792', 10),
    token: get('token') || 'dev',
  }).then((s) => {
    console.log('简阅服务已启动：' + s.url);
    console.log('扫描目录：\n  ' + s.roots().join('\n  '));
    if (args.includes('--open') && process.platform === 'win32') {
      try { require('child_process').exec('start "" "' + s.url + '"'); } catch (e) {}
    }
  });
}
