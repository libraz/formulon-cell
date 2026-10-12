/**
 * Compare two saved workbooks while tolerating the DOS timestamp that the ZIP
 * writer puts in each local and central-directory header.
 *
 * The comparison deliberately stays byte based.  A valid archive contributes
 * only header timestamp ranges to the ignore list; compressed data, checksums,
 * names, relationships, XML, and every other ZIP field still have to match.
 * If either input cannot be parsed as a bounded, structurally consistent ZIP
 * archive, the safe fallback is an exact byte comparison.
 */

const LOCAL_FILE_HEADER = 0x04034b50;
const CENTRAL_DIRECTORY_HEADER = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
const LOCAL_FILE_HEADER_SIZE = 30;
const CENTRAL_DIRECTORY_HEADER_SIZE = 46;
const END_OF_CENTRAL_DIRECTORY_SIZE = 22;
const MAX_ZIP_COMMENT_LENGTH = 0xffff;

interface ByteRange {
  readonly start: number;
  readonly end: number;
}

interface ParsedZip {
  readonly timestampRanges: readonly ByteRange[];
}

const readU16 = (bytes: Uint8Array, offset: number): number =>
  (bytes[offset] ?? 0) | ((bytes[offset + 1] ?? 0) << 8);

const readU32 = (bytes: Uint8Array, offset: number): number =>
  ((bytes[offset] ?? 0) |
    ((bytes[offset + 1] ?? 0) << 8) |
    ((bytes[offset + 2] ?? 0) << 16) |
    ((bytes[offset + 3] ?? 0) << 24)) >>>
  0;

const hasBytes = (bytes: Uint8Array, offset: number, length: number): boolean =>
  offset >= 0 && length >= 0 && offset <= bytes.length && length <= bytes.length - offset;

const equalBytes = (left: Uint8Array, right: Uint8Array): boolean => {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
};

const equalSlice = (
  left: Uint8Array,
  leftOffset: number,
  right: Uint8Array,
  rightOffset: number,
  length: number,
): boolean => {
  for (let index = 0; index < length; index += 1) {
    if (left[leftOffset + index] !== right[rightOffset + index]) return false;
  }
  return true;
};

const equalRanges = (left: readonly ByteRange[], right: readonly ByteRange[]): boolean => {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    const leftRange = left[index];
    const rightRange = right[index];
    if (
      leftRange === undefined ||
      rightRange === undefined ||
      leftRange.start !== rightRange.start ||
      leftRange.end !== rightRange.end
    ) {
      return false;
    }
  }
  return true;
};

/** Find the final EOCD whose comment reaches the end of the input. */
const findEndOfCentralDirectory = (bytes: Uint8Array): number | null => {
  if (bytes.length < END_OF_CENTRAL_DIRECTORY_SIZE) return null;
  const firstCandidate = Math.max(
    0,
    bytes.length - END_OF_CENTRAL_DIRECTORY_SIZE - MAX_ZIP_COMMENT_LENGTH,
  );
  for (
    let offset = bytes.length - END_OF_CENTRAL_DIRECTORY_SIZE;
    offset >= firstCandidate;
    offset -= 1
  ) {
    if (readU32(bytes, offset) !== END_OF_CENTRAL_DIRECTORY) continue;
    const commentLength = readU16(bytes, offset + 20);
    if (offset + END_OF_CENTRAL_DIRECTORY_SIZE + commentLength !== bytes.length) continue;
    const diskNumber = readU16(bytes, offset + 4);
    const centralDirectoryDisk = readU16(bytes, offset + 6);
    const entriesOnDisk = readU16(bytes, offset + 8);
    const entryCount = readU16(bytes, offset + 10);
    const centralDirectorySize = readU32(bytes, offset + 12);
    const centralDirectoryOffset = readU32(bytes, offset + 16);
    if (
      diskNumber !== 0 ||
      centralDirectoryDisk !== 0 ||
      entriesOnDisk !== entryCount ||
      entriesOnDisk === 0xffff ||
      entryCount === 0xffff ||
      centralDirectorySize === 0xffffffff ||
      centralDirectoryOffset === 0xffffffff ||
      centralDirectoryOffset > offset ||
      centralDirectorySize > offset - centralDirectoryOffset ||
      centralDirectoryOffset + centralDirectorySize !== offset
    ) {
      continue;
    }
    return offset;
  }
  return null;
};

/**
 * Parse the ordinary (non-ZIP64) ZIP shape emitted by the workbook writer.
 * Every local header is reached through a central-directory offset, so a
 * `PK` sequence in an entry name, comment, or payload cannot become a header.
 */
const parseZip = (bytes: Uint8Array): ParsedZip | null => {
  const endOfCentralDirectory = findEndOfCentralDirectory(bytes);
  if (endOfCentralDirectory === null) return null;

  const diskNumber = readU16(bytes, endOfCentralDirectory + 4);
  const centralDirectoryDisk = readU16(bytes, endOfCentralDirectory + 6);
  const entriesOnDisk = readU16(bytes, endOfCentralDirectory + 8);
  const entryCount = readU16(bytes, endOfCentralDirectory + 10);
  const centralDirectorySize = readU32(bytes, endOfCentralDirectory + 12);
  const centralDirectoryOffset = readU32(bytes, endOfCentralDirectory + 16);

  // ZIP64 values and multi-disk archives need a different parser. Falling
  // back to exact comparison is safer than masking an unknown field.
  if (
    diskNumber !== 0 ||
    centralDirectoryDisk !== 0 ||
    entriesOnDisk !== entryCount ||
    entriesOnDisk === 0xffff ||
    entryCount === 0xffff ||
    centralDirectorySize === 0xffffffff ||
    centralDirectoryOffset === 0xffffffff ||
    centralDirectoryOffset > endOfCentralDirectory ||
    centralDirectorySize > endOfCentralDirectory - centralDirectoryOffset
  ) {
    return null;
  }

  const centralDirectoryEnd = centralDirectoryOffset + centralDirectorySize;
  if (centralDirectoryEnd !== endOfCentralDirectory) return null;

  const timestampRanges: ByteRange[] = [];
  const localOffsets = new Set<number>();
  const occupiedRanges: ByteRange[] = [];
  let centralCursor = centralDirectoryOffset;

  for (let entry = 0; entry < entryCount; entry += 1) {
    if (
      !hasBytes(bytes, centralCursor, CENTRAL_DIRECTORY_HEADER_SIZE) ||
      readU32(bytes, centralCursor) !== CENTRAL_DIRECTORY_HEADER
    ) {
      return null;
    }

    const centralFlags = readU16(bytes, centralCursor + 8);
    const centralMethod = readU16(bytes, centralCursor + 10);
    const centralCrc = readU32(bytes, centralCursor + 16);
    const centralCompressedSize = readU32(bytes, centralCursor + 20);
    const centralUncompressedSize = readU32(bytes, centralCursor + 24);
    const centralNameLength = readU16(bytes, centralCursor + 28);
    const centralExtraLength = readU16(bytes, centralCursor + 30);
    const centralCommentLength = readU16(bytes, centralCursor + 32);
    const localOffset = readU32(bytes, centralCursor + 42);
    const centralHeaderEnd =
      centralCursor +
      CENTRAL_DIRECTORY_HEADER_SIZE +
      centralNameLength +
      centralExtraLength +
      centralCommentLength;
    if (centralHeaderEnd > centralDirectoryEnd) return null;

    if (
      localOffsets.has(localOffset) ||
      !hasBytes(bytes, localOffset, LOCAL_FILE_HEADER_SIZE) ||
      readU32(bytes, localOffset) !== LOCAL_FILE_HEADER
    ) {
      return null;
    }
    localOffsets.add(localOffset);

    const localFlags = readU16(bytes, localOffset + 6);
    const localMethod = readU16(bytes, localOffset + 8);
    const localCrc = readU32(bytes, localOffset + 14);
    const localCompressedSize = readU32(bytes, localOffset + 18);
    const localUncompressedSize = readU32(bytes, localOffset + 22);
    const localNameLength = readU16(bytes, localOffset + 26);
    const localExtraLength = readU16(bytes, localOffset + 28);
    const localHeaderEnd =
      localOffset + LOCAL_FILE_HEADER_SIZE + localNameLength + localExtraLength;
    const localDataEnd = localHeaderEnd + centralCompressedSize;

    if (
      localOffset >= centralDirectoryOffset ||
      localHeaderEnd > centralDirectoryOffset ||
      localFlags !== centralFlags ||
      localMethod !== centralMethod ||
      localNameLength !== centralNameLength ||
      !hasBytes(bytes, localOffset + LOCAL_FILE_HEADER_SIZE, localNameLength) ||
      !hasBytes(bytes, centralCursor + CENTRAL_DIRECTORY_HEADER_SIZE, centralNameLength) ||
      !equalSlice(
        bytes,
        localOffset + LOCAL_FILE_HEADER_SIZE,
        bytes,
        centralCursor + CENTRAL_DIRECTORY_HEADER_SIZE,
        centralNameLength,
      )
    ) {
      return null;
    }

    let localOccupiedEnd = localDataEnd;
    // A data-descriptor entry is allowed to leave CRC/sizes as zero in its
    // local header. Its 12/16 bytes are still part of the occupied local
    // record and must be validated before another central entry can point
    // into them.
    if ((centralFlags & 0x0008) !== 0) {
      const signedDescriptor =
        hasBytes(bytes, localDataEnd, 16) &&
        readU32(bytes, localDataEnd) === DATA_DESCRIPTOR_SIGNATURE &&
        readU32(bytes, localDataEnd + 4) === centralCrc &&
        readU32(bytes, localDataEnd + 8) === centralCompressedSize &&
        readU32(bytes, localDataEnd + 12) === centralUncompressedSize &&
        localDataEnd + 16 <= centralDirectoryOffset;
      const unsignedDescriptor =
        hasBytes(bytes, localDataEnd, 12) &&
        readU32(bytes, localDataEnd) === centralCrc &&
        readU32(bytes, localDataEnd + 4) === centralCompressedSize &&
        readU32(bytes, localDataEnd + 8) === centralUncompressedSize &&
        localDataEnd + 12 <= centralDirectoryOffset;
      if (signedDescriptor) {
        localOccupiedEnd = localDataEnd + 16;
      } else if (unsignedDescriptor) {
        localOccupiedEnd = localDataEnd + 12;
      } else {
        return null;
      }
    } else {
      if (
        localCrc !== centralCrc ||
        localCompressedSize !== centralCompressedSize ||
        localUncompressedSize !== centralUncompressedSize
      ) {
        return null;
      }
    }
    if (localOccupiedEnd > centralDirectoryOffset) return null;
    occupiedRanges.push({ start: localOffset, end: localOccupiedEnd });

    // DOS time (2 bytes) + DOS date (2 bytes), local then central. These are
    // the only fields intentionally ignored by sameSavedWorkbook.
    timestampRanges.push(
      { start: localOffset + 10, end: localOffset + 14 },
      { start: centralCursor + 12, end: centralCursor + 16 },
    );
    centralCursor = centralHeaderEnd;
  }

  if (centralCursor !== centralDirectoryEnd) return null;
  occupiedRanges.sort((left, right) => left.start - right.start);
  for (let index = 1; index < occupiedRanges.length; index += 1) {
    const previous = occupiedRanges[index - 1];
    const current = occupiedRanges[index];
    if (previous !== undefined && current !== undefined && current.start < previous.end) {
      return null;
    }
  }
  timestampRanges.sort((left, right) => left.start - right.start);
  return { timestampRanges };
};

const equalOutsideRanges = (
  left: Uint8Array,
  right: Uint8Array,
  ranges: readonly ByteRange[],
): boolean => {
  let rangeIndex = 0;
  for (let offset = 0; offset < left.length; offset += 1) {
    const range = ranges[rangeIndex];
    if (range !== undefined && offset >= range.end) rangeIndex += 1;
    const activeRange = ranges[rangeIndex];
    if (activeRange !== undefined && offset >= activeRange.start && offset < activeRange.end) {
      continue;
    }
    if (left[offset] !== right[offset]) return false;
  }
  return true;
};

/**
 * Return whether two saved workbook byte streams describe the same archive,
 * allowing only ZIP DOS timestamp fields to differ.
 */
export const sameSavedWorkbook = (left: Uint8Array, right: Uint8Array): boolean => {
  if (left.length !== right.length) return false;
  const leftZip = parseZip(left);
  const rightZip = parseZip(right);
  if (leftZip === null || rightZip === null) return equalBytes(left, right);
  if (!equalRanges(leftZip.timestampRanges, rightZip.timestampRanges)) return false;
  return equalOutsideRanges(left, right, leftZip.timestampRanges);
};
