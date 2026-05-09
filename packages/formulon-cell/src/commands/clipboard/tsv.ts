/**
 * Tab-separated parsing/encoding compatible with desktop spreadsheets.
 * RFC 4180 — but with `\t` as the delimiter and `\r\n` / `\n` as the row
 * terminator. Cells containing tabs, newlines or quotes are wrapped in
 * double quotes; embedded quotes are doubled.
 */

export function encodeTSV(rows: readonly (readonly string[])[]): string {
  return rows.map((row) => row.map(escapeCell).join('\t')).join('\r\n');
}

function escapeCell(cell: string): string {
  if (/[\t\r\n"]/.test(cell)) {
    return `"${cell.replace(/"/g, '""')}"`;
  }
  return cell;
}

export function parseTSV(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let i = 0;
  let quoted = false;
  // Strip a single trailing line terminator so we don't emit a phantom row.
  const src = text.replace(/(\r\n|\r|\n)$/, '');

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
    if (ch === '\t') {
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
