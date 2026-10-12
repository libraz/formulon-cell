import { describe, expect, it } from 'vitest';
import { sameSavedWorkbook } from '../../../src/commands/saved-workbook-proof.js';

interface ZipEntry {
  readonly name: string;
  readonly data: Uint8Array;
  readonly comment?: string;
  readonly method?: number;
  readonly flags?: number;
  readonly descriptorSignature?: boolean;
  readonly crc?: number;
  readonly time?: number;
  readonly date?: number;
}

interface BuiltZip {
  readonly bytes: Uint8Array;
  readonly timestampOffsets: readonly number[];
  readonly payloadOffsets: readonly number[];
  readonly crcOffsets: readonly number[];
  readonly nameOffsets: readonly number[];
  readonly localOffsets: readonly number[];
  readonly centralOffsets: readonly number[];
}

const encoder = new TextEncoder();

const writeU16 = (bytes: Uint8Array, offset: number, value: number): void => {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
};

const writeU32 = (bytes: Uint8Array, offset: number, value: number): void => {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
  bytes[offset + 2] = (value >>> 16) & 0xff;
  bytes[offset + 3] = (value >>> 24) & 0xff;
};

const concatenate = (parts: readonly Uint8Array[]): Uint8Array => {
  const bytes = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.length;
  }
  return bytes;
};

const makeZip = (entries: readonly ZipEntry[], archiveComment = ''): BuiltZip => {
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  const timestampOffsets: number[] = [];
  const payloadOffsets: number[] = [];
  const crcOffsets: number[] = [];
  const nameOffsets: number[] = [];
  const localOffsets: number[] = [];
  const centralOffsets: number[] = [];
  let localOffset = 0;

  for (const entry of entries) {
    const name = encoder.encode(entry.name);
    const comment = encoder.encode(entry.comment ?? '');
    const method = entry.method ?? 0;
    const flags = entry.flags ?? 0;
    const crc = entry.crc ?? 0x10203040;
    const time = entry.time ?? 0x1234;
    const date = entry.date ?? 0x5678;
    const hasDescriptor = (flags & 0x0008) !== 0;
    const descriptorLength = hasDescriptor ? (entry.descriptorSignature === false ? 12 : 16) : 0;
    const local = new Uint8Array(30 + name.length + entry.data.length + descriptorLength);
    writeU32(local, 0, 0x04034b50);
    writeU16(local, 4, 20);
    writeU16(local, 6, flags);
    writeU16(local, 8, method);
    writeU16(local, 10, time);
    writeU16(local, 12, date);
    if (!hasDescriptor) {
      writeU32(local, 14, crc);
      writeU32(local, 18, entry.data.length);
      writeU32(local, 22, entry.data.length);
    }
    writeU16(local, 26, name.length);
    writeU16(local, 28, 0);
    local.set(name, 30);
    local.set(entry.data, 30 + name.length);
    if (hasDescriptor) {
      const descriptorOffset = 30 + name.length + entry.data.length;
      if (descriptorLength === 16) {
        writeU32(local, descriptorOffset, 0x08074b50);
        writeU32(local, descriptorOffset + 4, crc);
        writeU32(local, descriptorOffset + 8, entry.data.length);
        writeU32(local, descriptorOffset + 12, entry.data.length);
      } else {
        writeU32(local, descriptorOffset, crc);
        writeU32(local, descriptorOffset + 4, entry.data.length);
        writeU32(local, descriptorOffset + 8, entry.data.length);
      }
    }
    timestampOffsets.push(localOffset + 10, localOffset + 12);
    payloadOffsets.push(localOffset + 30 + name.length);
    crcOffsets.push(localOffset + 14);
    nameOffsets.push(localOffset + 30);
    localOffsets.push(localOffset);
    locals.push(local);

    const centralOffset = localOffset;
    const central = new Uint8Array(46 + name.length + comment.length);
    writeU32(central, 0, 0x02014b50);
    writeU16(central, 4, 20);
    writeU16(central, 6, 20);
    writeU16(central, 8, flags);
    writeU16(central, 10, method);
    writeU16(central, 12, time);
    writeU16(central, 14, date);
    writeU32(central, 16, crc);
    writeU32(central, 20, entry.data.length);
    writeU32(central, 24, entry.data.length);
    writeU16(central, 28, name.length);
    writeU16(central, 30, 0);
    writeU16(central, 32, comment.length);
    writeU32(central, 42, centralOffset);
    central.set(name, 46);
    central.set(comment, 46 + name.length);
    centrals.push(central);
    localOffset += local.length;
  }

  const centralDirectoryOffset = localOffset;
  const centralDirectory = concatenate(centrals);
  const comment = encoder.encode(archiveComment);
  const end = new Uint8Array(22 + comment.length);
  writeU32(end, 0, 0x06054b50);
  writeU16(end, 8, entries.length);
  writeU16(end, 10, entries.length);
  writeU32(end, 12, centralDirectory.length);
  writeU32(end, 16, centralDirectoryOffset);
  writeU16(end, 20, comment.length);
  end.set(comment, 22);

  // The central records follow all local records, so derive their absolute
  // offsets from the known local-directory boundary after assembly.
  let centralCursor = centralDirectoryOffset;
  for (const central of centrals) {
    centralOffsets.push(centralCursor);
    timestampOffsets.push(centralCursor + 12, centralCursor + 14);
    crcOffsets.push(centralCursor + 16);
    nameOffsets.push(centralCursor + 46);
    centralCursor += central.length;
  }

  return {
    bytes: concatenate([...locals, centralDirectory, end]),
    timestampOffsets,
    payloadOffsets,
    crcOffsets,
    nameOffsets,
    localOffsets,
    centralOffsets,
  };
};

const copy = (bytes: Uint8Array): Uint8Array => new Uint8Array(bytes);

const writeTimestamp = (bytes: Uint8Array, offset: number, value: number): void =>
  writeU16(bytes, offset, value);

describe('sameSavedWorkbook', () => {
  it('ignores DOS timestamps in every local and central header', () => {
    const base = makeZip(
      [
        {
          name: 'xl/worksheets/PK\x03\x04-sheet.xml',
          data: encoder.encode('payload PK\x02\x01'),
          comment: 'entry comment PK\x03\x04',
          method: 8,
        },
        {
          name: 'xl/styles.xml',
          data: encoder.encode('styles'),
          comment: 'second entry',
        },
      ],
      'archive comment PK\x05\x06',
    );
    const later = copy(base.bytes);
    for (const [index, offset] of base.timestampOffsets.entries()) {
      writeTimestamp(later, offset, 0x7000 + index);
    }
    expect(sameSavedWorkbook(base.bytes, later)).toBe(true);
  });

  it('validates signed and unsigned data descriptors before masking timestamps', () => {
    const base = makeZip([
      {
        name: 'signed.bin',
        data: encoder.encode('signed payload'),
        flags: 0x0008,
      },
      {
        name: 'unsigned.bin',
        data: encoder.encode('unsigned payload'),
        flags: 0x0008,
        descriptorSignature: false,
      },
    ]);
    const later = copy(base.bytes);
    for (const [index, offset] of base.timestampOffsets.entries()) {
      writeTimestamp(later, offset, 0x2000 + index);
    }
    expect(sameSavedWorkbook(base.bytes, later)).toBe(true);
  });

  it('rejects a central entry whose local header overlaps another payload', () => {
    const base = makeZip([
      { name: 'first.bin', data: new Uint8Array(80) },
      { name: 'second.bin', data: new Uint8Array() },
    ]);
    const firstPayload = base.payloadOffsets[0];
    const secondLocal = base.localOffsets[1];
    const secondCentral = base.centralOffsets[1];
    if (firstPayload === undefined || secondLocal === undefined || secondCentral === undefined) {
      throw new Error('missing overlap offsets');
    }
    const fakeLocal = firstPayload + 10;
    const secondLocalLength = base.payloadOffsets[1];
    if (secondLocalLength === undefined) throw new Error('missing second payload offset');
    const malformed = copy(base.bytes);
    malformed.set(
      base.bytes.subarray(secondLocal, secondLocal + secondLocalLength - secondLocal),
      fakeLocal,
    );
    writeU32(malformed, secondCentral + 42, fakeLocal);

    const changed = copy(malformed);
    writeTimestamp(changed, fakeLocal + 10, 0x7100);
    writeTimestamp(changed, fakeLocal + 12, 0x6200);
    expect(sameSavedWorkbook(malformed, changed)).toBe(false);
  });

  it('falls back to exact comparison when a descriptor is forged as a local header', () => {
    const base = makeZip([
      { name: 'first.bin', data: encoder.encode('payload'), flags: 0x0008 },
      { name: 'second.bin', data: new Uint8Array() },
    ]);
    const firstPayload = base.payloadOffsets[0];
    const secondLocal = base.localOffsets[1];
    const secondCentral = base.centralOffsets[1];
    const secondPayload = base.payloadOffsets[1];
    if (
      firstPayload === undefined ||
      secondLocal === undefined ||
      secondCentral === undefined ||
      secondPayload === undefined
    ) {
      throw new Error('missing descriptor offsets');
    }
    const descriptorStart = firstPayload + encoder.encode('payload').length;
    const malformed = copy(base.bytes);
    malformed.set(base.bytes.subarray(secondLocal, secondPayload - secondLocal), descriptorStart);
    writeU32(malformed, secondCentral + 42, descriptorStart);

    const changed = copy(malformed);
    writeTimestamp(changed, descriptorStart + 10, 0x7200);
    writeTimestamp(changed, descriptorStart + 12, 0x6300);
    expect(sameSavedWorkbook(malformed, changed)).toBe(false);
  });

  it('detects payload, CRC, and metadata changes outside timestamp fields', () => {
    const base = makeZip([
      {
        name: 'xl/workbook.xml',
        data: encoder.encode('compressed payload'),
        method: 8,
      },
    ]);

    const payloadChanged = copy(base.bytes);
    const payloadOffset = base.payloadOffsets[0];
    if (payloadOffset === undefined) throw new Error('missing payload offset');
    payloadChanged[payloadOffset] = (payloadChanged[payloadOffset] ?? 0) ^ 0x01;
    expect(sameSavedWorkbook(base.bytes, payloadChanged)).toBe(false);

    const crcChanged = copy(base.bytes);
    const localCrcOffset = base.crcOffsets[0];
    const centralCrcOffset = base.crcOffsets[1];
    if (localCrcOffset === undefined || centralCrcOffset === undefined) {
      throw new Error('missing CRC offsets');
    }
    writeU32(crcChanged, localCrcOffset, 0x55667788);
    writeU32(crcChanged, centralCrcOffset, 0x55667788);
    expect(sameSavedWorkbook(base.bytes, crcChanged)).toBe(false);

    const nameChanged = copy(base.bytes);
    // The name has the same byte length, so this is a pure metadata change.
    const localNameOffset = base.nameOffsets[0];
    const centralNameOffset = base.nameOffsets[1];
    if (localNameOffset === undefined || centralNameOffset === undefined) {
      throw new Error('missing name offsets');
    }
    nameChanged[localNameOffset] = (nameChanged[localNameOffset] ?? 0) ^ 0x01;
    nameChanged[centralNameOffset] = (nameChanged[centralNameOffset] ?? 0) ^ 0x01;
    expect(sameSavedWorkbook(base.bytes, nameChanged)).toBe(false);
  });

  it('uses exact comparison when ZIP structure cannot be parsed', () => {
    const malformed = Uint8Array.from([0x50, 0x4b, 0x05, 0x06, 0x00]);
    expect(sameSavedWorkbook(malformed, copy(malformed))).toBe(true);

    const changed = copy(malformed);
    const lastByte = changed.length - 1;
    changed[lastByte] = (changed[lastByte] ?? 0) ^ 0x01;
    expect(sameSavedWorkbook(malformed, changed)).toBe(false);

    const valid = makeZip([{ name: 'entry', data: encoder.encode('data') }]);
    const broken = copy(valid.bytes);
    broken[0] = 0;
    for (const [index, offset] of valid.timestampOffsets.entries()) {
      writeTimestamp(broken, offset, 0x4000 + index);
    }
    expect(sameSavedWorkbook(valid.bytes, broken)).toBe(false);
  });
});
