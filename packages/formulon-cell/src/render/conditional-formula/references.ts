import { colFromLetters, MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { ParsedA1Range, ParsedRef } from './types.js';

function sheetNameMatchesIndex(name: string, sheetIndex: number): boolean {
  return name.trim().toLowerCase() === `sheet${sheetIndex + 1}`.toLowerCase();
}

function stripSupportedSheetQualifier(raw: string, sheetIndex: number): string | null {
  const body = raw.trim();
  if (!body.includes('!')) return body;
  if (body.startsWith("'")) {
    let name = '';
    for (let i = 1; i < body.length; i += 1) {
      const ch = body[i];
      if (ch === "'") {
        if (body[i + 1] === "'") {
          name += "'";
          i += 1;
          continue;
        }
        if (body[i + 1] !== '!') return null;
        return sheetNameMatchesIndex(name, sheetIndex) ? body.slice(i + 2).trim() : null;
      }
      name += ch;
    }
    return null;
  }
  const bang = body.indexOf('!');
  const sheetName = body.slice(0, bang);
  if (!/^[A-Za-z_][A-Za-z0-9_. ]*$/.test(sheetName)) return null;
  return sheetNameMatchesIndex(sheetName, sheetIndex) ? body.slice(bang + 1).trim() : null;
}

function parseA1Ref(raw: string, sheetIndex: number): ParsedRef | null {
  const body = stripSupportedSheetQualifier(raw, sheetIndex);
  if (body === null) return null;
  const m = body.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
  if (!m) return null;
  const col = colFromLetters(m[2] ?? '');
  const row = Number.parseInt(m[4] ?? '', 10) - 1;
  if (row < 0 || col < 0 || row > MAX_ROW || col > MAX_COL) return null;
  return { row, col, absCol: m[1] === '$', absRow: m[3] === '$' };
}

function parseR1C1Ref(
  raw: string,
  sheetIndex: number,
  baseRow: number,
  baseCol: number,
): ParsedRef | null {
  const body = stripSupportedSheetQualifier(raw, sheetIndex);
  if (body === null) return null;
  const m = body.match(/^R(?:(\d+)|\[([+-]?\d+)\])?C(?:(\d+)|\[([+-]?\d+)\])?$/i);
  if (!m) return null;
  const row =
    m[1] !== undefined ? Number.parseInt(m[1], 10) - 1 : baseRow + Number.parseInt(m[2] ?? '0', 10);
  const col =
    m[3] !== undefined ? Number.parseInt(m[3], 10) - 1 : baseCol + Number.parseInt(m[4] ?? '0', 10);
  if (row < 0 || col < 0 || row > MAX_ROW || col > MAX_COL) return null;
  return {
    row,
    col,
    absRow: m[2] === undefined,
    absCol: m[4] === undefined,
  };
}

function parseR1C1Range(
  raw: string,
  sheetIndex: number,
  baseRow: number,
  baseCol: number,
): ParsedA1Range | null {
  const body = stripSupportedSheetQualifier(raw, sheetIndex);
  if (body === null) return null;
  const parts = body.split(':');
  if (parts.length === 1) {
    const ref = parseR1C1Ref(parts[0] ?? '', sheetIndex, baseRow, baseCol);
    return ref ? { start: ref, end: ref } : null;
  }
  if (parts.length !== 2) return null;
  const start = parseR1C1Ref(parts[0] ?? '', sheetIndex, baseRow, baseCol);
  const end = parseR1C1Ref(parts[1] ?? '', sheetIndex, baseRow, baseCol);
  return start && end ? { start, end } : null;
}

function parseA1Range(raw: string, sheetIndex: number): ParsedA1Range | null {
  const body = stripSupportedSheetQualifier(raw, sheetIndex);
  if (body === null) return null;
  const parts = body.split(':');
  if (parts.length === 1) {
    const ref = parseA1Ref(parts[0] ?? '', sheetIndex);
    return ref ? { start: ref, end: ref } : null;
  }
  if (parts.length !== 2) return null;
  const start = parseA1Ref(parts[0] ?? '', sheetIndex);
  const end = parseA1Ref(parts[1] ?? '', sheetIndex);
  return start && end ? { start, end } : null;
}

export { parseA1Range, parseA1Ref, parseR1C1Range, parseR1C1Ref };
