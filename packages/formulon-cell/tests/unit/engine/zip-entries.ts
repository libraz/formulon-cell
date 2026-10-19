import { inflateRawSync } from 'node:zlib';

/** Every entry of a zip package by path, inflated. Enough for the stored and
 *  deflated entries an xlsx holds; no zip64, no encryption. */
export function zipEntries(bytes: Uint8Array): Map<string, Uint8Array> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.length - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06054b50) end -= 1;
  if (end < 0) throw new Error('zip: no end of central directory');
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const entries = new Map<string, Uint8Array>();
  const decoder = new TextDecoder();
  for (let i = 0; i < count; i += 1) {
    if (view.getUint32(at, true) !== 0x02014b50) throw new Error('zip: bad central entry');
    const method = view.getUint16(at + 10, true);
    const compressed = view.getUint32(at + 20, true);
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const local = view.getUint32(at + 42, true);
    const name = decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength));
    const dataStart =
      local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const raw = bytes.subarray(dataStart, dataStart + compressed);
    entries.set(name, method === 0 ? raw.slice() : new Uint8Array(inflateRawSync(raw)));
    at += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/** The text of a package entry, or '' when the package has none. */
export function zipText(entries: Map<string, Uint8Array>, path: string): string {
  const entry = entries.get(path);
  return entry ? new TextDecoder().decode(entry) : '';
}
