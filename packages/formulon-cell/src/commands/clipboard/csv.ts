/**
 * CSV parsing/encoding with an explicit comma or semicolon separator.
 * Fields containing the selected separator, newlines or quotes are wrapped
 * in double quotes; embedded quotes are doubled. Comma is the default.
 */

const BOM = '﻿';

/** Separators supported by the CSV document format. */
export type CSVDelimiter = ',' | ';';

export interface CSVParseOptions {
  /** Field separator. Comma remains the default for backwards compatibility. */
  delimiter?: CSVDelimiter;
}

export interface CSVEncodeOptions {
  /** Field separator. Comma remains the default for backwards compatibility. */
  delimiter?: CSVDelimiter;
  /** Use \n instead of \r\n. Spreadsheets write \r\n; reading is robust to either. */
  eol?: '\r\n' | '\n';
  /** Prepend a UTF-8 BOM. Desktop spreadsheets on Windows expects this for correct UTF-8
   *  detection when opening a .csv via double-click. */
  bom?: boolean;
}

const resolveDelimiter = (delimiter: CSVDelimiter | undefined): CSVDelimiter => {
  if (delimiter === undefined || delimiter === ',' || delimiter === ';') {
    return delimiter ?? ',';
  }
  // The public type protects TypeScript callers, but the runtime check keeps
  // JavaScript callers from silently producing a file with an unsupported
  // separator.
  throw new TypeError('CSV delimiter must be "," or ";"');
};

export function encodeCSV(
  rows: readonly (readonly string[])[],
  opts: CSVEncodeOptions = {},
): string {
  const delimiter = resolveDelimiter(opts.delimiter);
  const eol = opts.eol ?? '\r\n';
  const body = rows
    .map((row) => row.map((cell) => escapeCell(cell, delimiter)).join(delimiter))
    .join(eol);
  return opts.bom ? BOM + body : body;
}

function escapeCell(cell: string, delimiter: CSVDelimiter): string {
  // Quote whenever the cell would be ambiguous — separators, line breaks, quote
  // chars, or leading/trailing whitespace (spreadsheets preserve whitespace inside
  // quoted fields but trims unquoted ones on read).
  if (cell.includes(delimiter) || /[\r\n"]|^\s|\s$/.test(cell)) {
    return `"${cell.replace(/"/g, '""')}"`;
  }
  return cell;
}

export function parseCSV(text: string, opts: CSVParseOptions = {}): string[][] {
  const delimiter = resolveDelimiter(opts.delimiter);
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
    if (ch === delimiter) {
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
