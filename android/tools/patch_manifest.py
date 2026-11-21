#!/usr/bin/env python3
"""就地改写二进制 AndroidManifest.xml 的 versionCode / versionName（等长替换）。"""
import struct, sys

def read_str(pool, idx):
    offs = pool['offsets']
    raw = pool['data']
    o = pool['strings_start'] + offs[idx]
    utf8 = pool['utf8']
    if utf8:
        n = raw[o]
        if n & 0x80:
            n = ((n & 0x7f) << 8) | raw[o + 1]
            o += 2
        else:
            o += 1
        return raw[o:o + n].decode('utf-8', 'replace'), o, n
    n = struct.unpack_from('<H', raw, o)[0]
    if n & 0x8000:
        n = ((n & 0x7fff) << 16) | struct.unpack_from('<H', raw, o + 2)[0]
        o += 4
    else:
        o += 2
    return raw[o:o + n * 2].decode('utf-16le', 'replace'), o, n

def main(src, dst, version_code, version_name):
    data = bytearray(open(src, 'rb').read())
    if struct.unpack_from('<H', data, 0)[0] != 0x0003:
        raise SystemExit('不是 AXML')
    pos = 8
    pool = None
    resmap = []
    while pos + 8 <= len(data):
        ctype, hsize, csize = struct.unpack_from('<HHI', data, pos)
        if ctype == 0x0001:
            sc, styc, flags, sstart, styoff = struct.unpack_from('<IIIII', data, pos + 8)
            offs = list(struct.unpack_from('<%dI' % sc, data, pos + 28))
            pool = {'data': data, 'offsets': offs,
                    'utf8': bool(flags & 0x100), 'strings_start': pos + sstart, 'count': sc}
        elif ctype == 0x0180:
            n = (csize - hsize) // 4
            resmap = list(struct.unpack_from('<%dI' % n, data, pos + hsize))
        elif ctype == 0x0102 and pool:
            p = pos + hsize
            ns, name, attr_start, attr_size, attr_count, id_index, class_index, style_index = struct.unpack_from('<IIHHHHHH', data, p)
            elname, _, _ = read_str(pool, name)
            if elname == 'manifest':
                fixed = False
                for i in range(attr_count):
                    a = p + attr_start + i * attr_size
                    ans, aname, araw, tsize, res0, dtype, aval = struct.unpack_from('<IIIHBBI', data, a)
                    nm, _, _ = read_str(pool, aname)
                    rid = resmap[aname] if aname < len(resmap) else 0
                    if rid == 0x0101021b:
                        struct.pack_into('<I', data, a + 16, int(version_code))
                        fixed = True
                        print('versionCode %s → %s' % (aval, version_code))
                    elif rid == 0x0101021c:
                        # 找到 rawValue 指向的字符串，等长替换
                        si = araw
                        old, off, n = read_str(pool, si)
                        if len(old) != len(version_name):
                            raise SystemExit('versionName 长度必须一致：%r → %r' % (old, version_name))
                        enc = version_name.encode('utf-8' if pool['utf8'] else 'utf-16le')
                        data[off:off + len(enc)] = enc
                        fixed = True
                        print('versionName %s → %s' % (old, version_name))
                    elif nm == 'versionCode' and rid == 0:
                        struct.pack_into('<I', data, a + 16, int(version_code))
                if fixed:
                    break
        pos += csize
    open(dst, 'wb').write(data)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
