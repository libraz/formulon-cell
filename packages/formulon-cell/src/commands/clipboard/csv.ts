/**
 * Comma-separated parsing/encoding compatible with desktop spreadsheets.
 * RFC 4180 — fields containing commas, newlines or quotes are wrapped in
 * double quotes; embedded quotes are doubled.
 *
 * Mirrors `tsv.ts` with `,` as the delimiter. Kept as a sibling rather than a
 * shared parameterised parser because the encoding edge cases (locale-style
 * decimal commas in non-US desktop spreadsheets exports) are CSV-specific and we'd rather
 * iterate on them in one file.
 */

const BOM = '﻿';

export interface CSVEncodeOptions {
  /** Use \n instead of \r\n. Spreadsheets write \r\n; reading is robust to either. */
  eol?: '\r\n' | '\n';
  /** Prepend a UTF-8 BOM. Desktop spreadsheets on Windows expects this for correct UTF-8
   *  detection when opening a .csv via double-click. */
  bom?: boolean;
}

export function encodeCSV(
  rows: readonly (readonly string[])[],
  opts: CSVEncodeOptions = {},
): string {
  const eol = opts.eol ?? '\r\n';
  const body = rows.map((row) => row.map(escapeCell).join(',')).join(eol);
  return opts.bom ? BOM + body : body;
}

function escapeCell(cell: string): string {
  // Quote whenever the cell would be ambiguous — commas, line breaks, quote
  // chars, or leading/trailing whitespace (spreadsheets preserve whitespace inside
  // quoted fields but trims unquoted ones on read).
  if (/[,\r\n"]|^\s|\s$/.test(cell)) {
    return `"${cell.replace(/"/g, '""')}"`;
  }
  return cell;
}

export function parseCSV(text: string): string[][] {
  let src = text;
  // Strip optional UTF-8 BOM.
  if (src.charCodeAt(0) === 0xfeff) src = src.slice(1);
  // Strip a single trailing line terminator so we don't emit a phantom row.
  src = src.replace(/(\r\n|\r|\n)$/, '');

  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let i = 0;
  let quoted = false;

  while (i < src.length) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          cell += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i += 1;
        continue;
      }
      cell += ch;
      i += 1;
      continue;
    }
    if (ch === '"' && cell === '') {
      quoted = true;
      i += 1;
      continue;
    }
    if (ch === ',') {
      row.push(cell);
      cell = '';
      i += 1;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      // Swallow \r\n as a single break.
      if (ch === '\r' && src[i + 1] === '\n') i += 2;
      else i += 1;
      continue;
    }
    cell += ch;
    i += 1;
  }
  // Final cell (no trailing terminator).
  row.push(cell);
  rows.push(row);
  return rows;
}
