import { crc32, deflateRawSync, inflateRawSync } from 'node:zlib';

/** One member of a zip archive, with its decompressed bytes. */
export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

const u16 = (b: Uint8Array, o: number): number => (b[o] ?? 0) | ((b[o + 1] ?? 0) << 8);
const u32 = (b: Uint8Array, o: number): number => u16(b, o) + u16(b, o + 2) * 0x10000;

/** Reads every member of a zip archive via its central directory. */
export function readZip(bytes: Uint8Array): ZipEntry[] {
  let eocd = bytes.length - 22;
  while (eocd >= 0 && u32(bytes, eocd) !== 0x06054b50) eocd--;
  if (eocd < 0) throw new Error('not a zip archive');
  const count = u16(bytes, eocd + 10);
  let p = u32(bytes, eocd + 16);
  const out: ZipEntry[] = [];
  for (let i = 0; i < count; i++) {
    const method = u16(bytes, p + 10);
    const size = u32(bytes, p + 20);
    const nameLen = u16(bytes, p + 28);
    const extraLen = u16(bytes, p + 30);
    const commentLen = u16(bytes, p + 32);
    const local = u32(bytes, p + 42);
    const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLen));
    const dataStart = local + 30 + u16(bytes, local + 26) + u16(bytes, local + 28);
    const raw = bytes.subarray(dataStart, dataStart + size);
    out.push({ name, data: method === 0 ? raw.slice() : new Uint8Array(inflateRawSync(raw)) });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

/** Writes a zip archive with deflated members. */
export function writeZip(entries: readonly ZipEntry[]): Uint8Array {
  const parts: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const enc = new TextEncoder();
  for (const { name, data } of entries) {
    const nameBytes = enc.encode(name);
    const packed = new Uint8Array(deflateRawSync(data));
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(8, 8, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, packed.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, nameBytes.length, true);
    const head = new DataView(new ArrayBuffer(46));
    head.setUint32(0, 0x02014b50, true);
    head.setUint16(4, 20, true);
    head.setUint16(6, 20, true);
    head.setUint16(10, 8, true);
    head.setUint32(16, crc, true);
    head.setUint32(20, packed.length, true);
    head.setUint32(24, data.length, true);
    head.setUint16(28, nameBytes.length, true);
    head.setUint32(42, offset, true);
    parts.push(new Uint8Array(local.buffer), nameBytes, packed);
    central.push(new Uint8Array(head.buffer), nameBytes);
    offset += 30 + nameBytes.length + packed.length;
  }
  const cdSize = central.reduce((n, c) => n + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, cdSize, true);
  end.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(offset + cdSize + 22);
  let at = 0;
  for (const part of all) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/** Rewrites the text of the member at `path` through `edit`; throws when the member is absent. */
export function patchZipText(
  bytes: Uint8Array,
  path: string,
  edit: (text: string) => string,
): Uint8Array {
  const entries = readZip(bytes);
  const target = entries.find((e) => e.name === path);
  if (!target) throw new Error(`missing zip member ${path}`);
  target.data = new TextEncoder().encode(edit(new TextDecoder().decode(target.data)));
  return writeZip(entries);
}
