/**
 * Desktop-spreadsheet-compatible HTML clipboard encoder. The output is a `<table>`
 * with inline styles for bold/italic/underline/strike/align/color/fill —
 * Spreadsheets parse these on paste.
 */

import { addrKey } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import { formatCell } from '../../engine/value.js';
import type { CellFormat, State } from '../../store/store.js';
import { formatNumber } from '../format.js';

const escapeHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const MAX_HTML_CLIPBOARD_CELLS = 1_000_000;

const styleOf = (fmt: CellFormat | undefined): string => {
  if (!fmt) return '';
  const parts: string[] = [];
  if (fmt.bold) parts.push('font-weight:bold');
  if (fmt.italic) parts.push('font-style:italic');
  const decos: string[] = [];
  if (fmt.underline) decos.push('underline');
  if (fmt.strike) decos.push('line-through');
  if (decos.length > 0) parts.push(`text-decoration:${decos.join(' ')}`);
  if (fmt.color) parts.push(`color:${fmt.color}`);
  if (fmt.fill) parts.push(`background-color:${fmt.fill}`);
  if (fmt.align) parts.push(`text-align:${fmt.align}`);
  if (fmt.fontFamily) parts.push(`font-family:${fmt.fontFamily}`);
  if (fmt.fontSize) parts.push(`font-size:${fmt.fontSize}px`);
  return parts.join(';');
};

/** Render the range as an HTML `<table>` with inline styles. */
export function encodeHtml(state: State, range: Range): string {
  const rowsCount = range.r1 - range.r0 + 1;
  const colsCount = range.c1 - range.c0 + 1;
  if (rowsCount <= 0 || colsCount <= 0) return '<table></table>';
  if (rowsCount * colsCount > MAX_HTML_CLIPBOARD_CELLS) return '';

  const rows: string[] = [];
  for (let r = range.r0; r <= range.r1; r += 1) {
    const cells: string[] = [];
    for (let c = range.c0; c <= range.c1; c += 1) {
      const key = addrKey({ sheet: range.sheet, row: r, col: c });
      const cell = state.data.cells.get(key);
      const fmt = state.format.formats.get(key);
      // Formula cells emit their formula text verbatim (`=...`); Excel and
      //  Sheets parse `=`-prefixed cell content on paste and rebuild the
      //  formula rather than freezing the last computed value.
      const text = !cell
        ? (fmt?.hyperlinkDisplay ?? '')
        : cell.formula != null
          ? cell.formula
          : cell.value.kind === 'blank'
            ? (fmt?.hyperlinkDisplay ?? '')
            : cell.value.kind === 'number' && fmt?.numFmt
              ? formatNumber(cell.value.value, fmt.numFmt)
              : formatCell(cell.value);
      const style = styleOf(fmt);
      const styleAttr = style ? ` style="${style}"` : '';
      const body = fmt?.hyperlink
        ? `<a href="${escapeHtml(fmt.hyperlink)}">${escapeHtml(text)}</a>`
        : escapeHtml(text);
      cells.push(`<td${styleAttr}>${body}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }
  return `<table>${rows.join('')}</table>`;
}
