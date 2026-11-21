/* OOXML 压缩包读取 + 媒体资源解析 */
import { IMG_EXT, mimeOf, bytesToDataUrl } from './oxml.js';

export class OoxmlZip {
  constructor(bytes) {
    const fflate = globalThis.fflate;
    this.files = fflate.unzipSync(new Uint8Array(bytes), { filter: (f) => f.size > 0 });
    this.lower = new Map();
    for (const k of Object.keys(this.files)) this.lower.set(k.toLowerCase(), k);
    this._urls = new Map();
  }
  has(name) { return this.lower.has(String(name).toLowerCase()); }
  get(name) {
    const k = this.lower.get(String(name).toLowerCase());
    return k ? this.files[k] : null;
  }
  text(name) {
    const b = this.get(name);
    if (!b) return null;
    return new TextDecoder('utf-8').decode(b);
  }
  url(name) {
    const key = String(name).toLowerCase();
    if (this._urls.has(key)) return this._urls.get(key);
    const b = this.get(name);
    if (!b) return null;
    let u = null;
    if (IMG_EXT.test(name)) {
      if (typeof Blob !== 'undefined' && typeof URL !== 'undefined' && URL.createObjectURL) {
        try { u = URL.createObjectURL(new Blob([b], { type: mimeOf(name) })); } catch (e) { u = null; }
      }
      if (!u) u = bytesToDataUrl(b, mimeOf(name));
    }
    this._urls.set(key, u);
    return u;
  }
  revoke() {
    if (typeof URL !== 'undefined' && URL.revokeObjectURL) {
      for (const u of this._urls.values()) if (u && u.startsWith('blob:')) URL.revokeObjectURL(u);
    }
    this._urls.clear();
  }
}

/** 解析 .rels 成 {id: target} */
export function parseRels(xmlTree) {
  const map = {};
  if (!xmlTree) return map;
  const list = xmlTree.getElementsByTagNameNS('*', 'Relationship');
  for (const r of list) {
    const id = r.getAttribute('Id');
    let target = r.getAttribute('Target') || '';
    const abs = target.startsWith('/');
    if (abs) target = target.slice(1);
    map[id] = { target, abs, type: r.getAttribute('Type') || '', external: r.getAttribute('TargetMode') === 'External' };
  }
  return map;
}

/** 把 rels 中的相对路径解析成 zip 内绝对路径 */
export function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (baseDir ? baseDir.split('/') : []).concat(target.split('/'));
  const out = [];
  for (const p of parts) {
    if (p === '' || p === '.') continue;
    if (p === '..') out.pop(); else out.push(p);
  }
  return out.join('/');
}

/** 解析关系目标为 zip 内绝对路径（正确处理 /开头的绝对路径） */
export function resolveRel(baseDir, rel) {
  if (!rel) return null;
  if (rel.abs) return rel.target;
  return resolvePath(baseDir, rel.target);
}
