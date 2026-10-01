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

  // HTML tables represent Excel's merged cells with rowspan/colspan. Keep
  // only merges fully contained by the materialized range; a partial source
  // merge cannot be represented without changing the selected rectangle.
  const mergeByAnchor = new Map<string, Range>();
  const coveredByMerge = new Set<string>();
  for (const merge of state.merges.byAnchor.values()) {
    if (
      merge.sheet !== range.sheet ||
      merge.r0 < range.r0 ||
      merge.c0 < range.c0 ||
      merge.r1 > range.r1 ||
      merge.c1 > range.c1
    ) {
      continue;
    }
    const anchorKey = addrKey({ sheet: merge.sheet, row: merge.r0, col: merge.c0 });
    mergeByAnchor.set(anchorKey, merge);
    for (let r = merge.r0; r <= merge.r1; r += 1) {
      for (let c = merge.c0; c <= merge.c1; c += 1) {
        if (r === merge.r0 && c === merge.c0) continue;
        coveredByMerge.add(addrKey({ sheet: merge.sheet, row: r, col: c }));
      }
    }
  }

  const rows: string[] = [];
  for (let r = range.r0; r <= range.r1; r += 1) {
    const cells: string[] = [];
    for (let c = range.c0; c <= range.c1; c += 1) {
      const key = addrKey({ sheet: range.sheet, row: r, col: c });
      if (coveredByMerge.has(key)) continue;
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
      const merge = mergeByAnchor.get(key);
      const spanAttrs = merge
        ? ` rowspan="${merge.r1 - merge.r0 + 1}" colspan="${merge.c1 - merge.c0 + 1}"`
        : '';
      cells.push(`<td${spanAttrs}${styleAttr}>${body}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }
  return `<table>${rows.join('')}</table>`;
}
