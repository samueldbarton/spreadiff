/**
 * Compound File Binary (OLE2) reader: the container of legacy .xls files.
 * Returns the streams at the top of the storage tree by name. Read-only,
 * no validation beyond what's needed not to loop forever.
 */

const FREE = 0xffffffff;
const END = 0xfffffffe;

/** Streams of a compound file, by name (`Workbook`, `Book`, …). */
export function readCfb(bytes: Uint8Array): Map<string, Uint8Array> {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (dv.getUint32(0, true) !== 0xe011cfd0 || dv.getUint32(4, true) !== 0xe11ab1a1) throw new Error('Not a compound file');
  const sectorSize = 1 << dv.getUint16(30, true);
  const miniSize = 1 << dv.getUint16(32, true);
  const dirStart = dv.getUint32(48, true);
  const miniCutoff = dv.getUint32(56, true);
  const miniFatStart = dv.getUint32(60, true);
  let difatStart = dv.getUint32(68, true);
  const difatCount = dv.getUint32(72, true);
  const sector = (i: number) => {
    const off = (i + 1) * sectorSize;
    return bytes.subarray(off, Math.min(off + sectorSize, bytes.length));
  };
  // The FAT's sectors: 109 listed in the header, the rest in a chain of DIFAT sectors.
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i++) {
    const s = dv.getUint32(76 + i * 4, true);
    if (s !== FREE && s !== END) fatSectors.push(s);
  }
  for (let n = 0; n < difatCount && difatStart !== END && difatStart !== FREE; n++) {
    const s = sector(difatStart);
    const sdv = new DataView(s.buffer, s.byteOffset, s.byteLength);
    const per = sectorSize / 4 - 1;
    for (let i = 0; i < per; i++) {
      const x = sdv.getUint32(i * 4, true);
      if (x !== FREE && x !== END) fatSectors.push(x);
    }
    difatStart = sdv.getUint32(per * 4, true);
  }
  const fat = new Uint32Array(fatSectors.length * (sectorSize / 4));
  fatSectors.forEach((s, k) => {
    const sec = sector(s);
    fat.set(new Uint32Array(sec.buffer.slice(sec.byteOffset, sec.byteOffset + sectorSize)), k * (sectorSize / 4));
  });
  const chain = (start: number, table: Uint32Array, limit: number): number[] => {
    const out: number[] = [];
    for (let s = start; s !== END && s !== FREE && s < table.length && out.length <= limit; s = table[s]) out.push(s);
    return out;
  };
  const maxSectors = Math.ceil(bytes.length / sectorSize) + 1;
  const read = (start: number, size: number): Uint8Array => {
    const out = new Uint8Array(size);
    let pos = 0;
    for (const s of chain(start, fat, maxSectors)) {
      if (pos >= size) break;
      const sec = sector(s);
      out.set(sec.subarray(0, Math.min(sec.length, size - pos)), pos);
      pos += sec.length;
    }
    return out;
  };
  // Directory: 128-byte entries.
  const dirSectors = chain(dirStart, fat, maxSectors);
  const dir = read(dirStart, dirSectors.length * sectorSize);
  const ddv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const entries: { name: string; type: number; start: number; size: number; left: number; right: number; child: number }[] = [];
  for (let off = 0; off + 128 <= dir.length; off += 128) {
    const nameLen = ddv.getUint16(off + 64, true);
    let name = '';
    for (let i = 0; i < Math.max(0, nameLen / 2 - 1) && i < 32; i++) name += String.fromCharCode(ddv.getUint16(off + i * 2, true));
    entries.push({
      name,
      type: dir[off + 66],
      left: ddv.getUint32(off + 68, true),
      right: ddv.getUint32(off + 72, true),
      child: ddv.getUint32(off + 76, true),
      start: ddv.getUint32(off + 116, true),
      size: ddv.getUint32(off + 120, true),
    });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new Error('Compound file without a root entry');
  // Small streams live in the mini stream (the root's data), in mini sectors.
  const miniStream = read(root.start, root.size);
  const miniFatSectors = chain(miniFatStart, fat, maxSectors);
  const miniFatBytes = read(miniFatStart, miniFatSectors.length * sectorSize);
  const miniFat = new Uint32Array(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength / 4);
  const readMini = (start: number, size: number): Uint8Array => {
    const out = new Uint8Array(size);
    let pos = 0;
    for (const s of chain(start, miniFat, Math.ceil(miniStream.length / miniSize) + 1)) {
      if (pos >= size) break;
      const off = s * miniSize;
      out.set(miniStream.subarray(off, Math.min(off + miniSize, off + size - pos)), pos);
      pos += miniSize;
    }
    return out;
  };
  // The root's children: a red-black tree, walked through left/right/child links.
  const streams = new Map<string, Uint8Array>();
  const seen = new Set<number>();
  const walk = (i: number) => {
    if (i === FREE || i >= entries.length || seen.has(i)) return;
    seen.add(i);
    const e = entries[i];
    walk(e.left);
    walk(e.right);
    if (e.type === 2) streams.set(e.name, e.size < miniCutoff ? readMini(e.start, e.size) : read(e.start, e.size));
  };
  walk(root.child);
  return streams;
}
