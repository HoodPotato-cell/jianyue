/* OLE2 / Compound File Binary 读取器（只读，够 .doc/.xls/.ppt 用） */

const END = 0xFFFFFFFA;

export function isCfb(buf) {
  const u8 = new Uint8Array(buf);
  return u8.length > 512 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0;
}

export function parseCfb(buf) {
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  if (dv.getUint32(0, true) !== 0xe011cfd0 || dv.getUint32(4, true) !== 0xe11ab1a1) {
    throw new Error('不是 OLE2 复合文档');
  }
  const secShift = dv.getUint16(30, true) || 9;
  const secSize = 1 << secShift;
  const miniSize = 1 << (dv.getUint16(32, true) || 6);
  const nFat = dv.getUint32(44, true);
  const dirStart = dv.getUint32(48, true);
  const miniCutoff = dv.getUint32(56, true) || 4096;
  const miniFatStart = dv.getUint32(60, true);
  const nMiniFat = dv.getUint32(64, true);
  const difatStart = dv.getUint32(68, true);
  const nDifat = dv.getUint32(72, true);

  const secOff = (s) => (s + 1) * secSize;
  const inRange = (s) => s >= 0 && secOff(s) + secSize <= u8.length;

  // DIFAT → FAT 扇区表
  const fatSecs = [];
  for (let i = 0; i < 109 && fatSecs.length < nFat; i++) {
    const v = dv.getUint32(76 + i * 4, true);
    if (v < END && inRange(v)) fatSecs.push(v);
  }
  let d = difatStart;
  const per = secSize / 4 - 1;
  for (let k = 0; k < nDifat && d < END && inRange(d); k++) {
    for (let i = 0; i < per && fatSecs.length < nFat; i++) {
      const v = dv.getUint32(secOff(d) + i * 4, true);
      if (v < END && inRange(v)) fatSecs.push(v);
    }
    d = dv.getUint32(secOff(d) + per * 4, true);
  }

  const fat = new Int32Array(fatSecs.length * (secSize / 4));
  fatSecs.forEach((s, i) => {
    const base = i * (secSize / 4);
    for (let j = 0; j < secSize / 4; j++) fat[base + j] = dv.getInt32(secOff(s) + j * 4, true);
  });

  function chain(start, table, limit) {
    const out = [];
    let s = start;
    let guard = 0;
    while (s >= 0 && s < (limit || table.length) && s < END && guard++ < 1e6) {
      out.push(s);
      const nxt = table[s];
      if (nxt < 0 || nxt >= END || nxt === s) break;
      s = nxt;
    }
    return out;
  }

  function readSectors(start, size) {
    const secs = chain(start, fat, fat.length);
    const out = new Uint8Array(secs.length * secSize);
    secs.forEach((s, i) => {
      const o = secOff(s);
      out.set(u8.subarray(o, Math.min(o + secSize, u8.length)), i * secSize);
    });
    return size ? out.subarray(0, Math.min(size, out.length)) : out;
  }

  // 目录
  const dirBuf = readSectors(dirStart);
  const entries = [];
  for (let off = 0; off + 128 <= dirBuf.length; off += 128) {
    const nameLen = dirBuf[off + 64] | (dirBuf[off + 65] << 8);
    if (!nameLen) { entries.push(null); continue; }
    let name = '';
    for (let i = 0; i + 1 < nameLen && i < 64; i += 2) {
      const c = dirBuf[off + i] | (dirBuf[off + i + 1] << 8);
      if (!c) break;
      name += String.fromCharCode(c);
    }
    const type = dirBuf[off + 66];
    const sStart = (dirBuf[off + 116] | (dirBuf[off + 117] << 8) | (dirBuf[off + 118] << 16) | (dirBuf[off + 119] << 24)) >>> 0;
    const size = (dirBuf[off + 120] | (dirBuf[off + 121] << 8) | (dirBuf[off + 122] << 16) | (dirBuf[off + 123] << 24)) >>> 0;
    entries.push({ name, type, start: sStart, size });
  }
  const root = entries.find((e) => e && e.type === 5);
  if (!root) throw new Error('OLE2 缺少根目录');

  const miniStream = readSectors(root.start, root.size);
  let miniFatTable = null;
  if (miniFatStart < END) {
    const mf = readSectors(miniFatStart, nMiniFat * secSize);
    miniFatTable = new Int32Array(mf.buffer, mf.byteOffset, Math.floor(mf.byteLength / 4));
  }

  function readMini(start, size) {
    if (!miniFatTable) return new Uint8Array(0);
    const secs = chain(start, miniFatTable, miniFatTable.length);
    const out = new Uint8Array(secs.length * miniSize);
    secs.forEach((s, i) => {
      const o = s * miniSize;
      out.set(miniStream.subarray(o, Math.min(o + miniSize, miniStream.length)), i * miniSize);
    });
    return size ? out.subarray(0, Math.min(size, out.length)) : out;
  }

  const byName = new Map();
  for (const e of entries) {
    if (e && e.type === 2) byName.set(e.name.toLowerCase(), e);
  }

  return {
    entries: entries.filter((e) => e && e.type === 2),
    has(name) { return byName.has(String(name).toLowerCase()); },
    read(name) {
      const e = byName.get(String(name).toLowerCase());
      if (!e) throw new Error('缺少流：' + name);
      return e.size < miniCutoff && e.size > 0 ? readMini(e.start, e.size) : readSectors(e.start, e.size);
    },
    names() { return entries.filter((e) => e && e.type === 2).map((e) => e.name); },
  };
}
