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
import type { ClipboardFormatResolver } from './snapshot.js';

const escapeHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const splitFontFamily = (family: string): string[] => {
  const parts: string[] = [];
  let start = 0;
  let quote = '';
  let depth = 0;
  for (let i = 0; i < family.length; i += 1) {
    const ch = family[i] ?? '';
    if (quote) {
      if (ch === quote && family[i - 1] !== '\\') quote = '';
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') depth += 1;
    else if (ch === ')' && depth > 0) depth -= 1;
    else if (ch === ',' && depth === 0) {
      parts.push(family.slice(start, i).trim());
      start = i + 1;
    }
  }
  parts.push(family.slice(start).trim());
  return parts.filter(Boolean);
};

/** Keep a family name inside one CSS token before the style attribute is
 * HTML-escaped. This preserves ordinary stacks (`Arial, sans-serif`) while
 * making quotes, angle brackets, and declaration separators data. */
const cssString = (value: string): string =>
  `"${value
    .replaceAll('\\', '\\\\')
    .replaceAll('"', '\\"')
    .replace(/\r\n?|\n/g, '\\A ')
    .replaceAll('<', '\\3C ')
    .replaceAll('>', '\\3E ')}"`;

const safeFontFamily = (family: string): string =>
  splitFontFamily(family)
    .map((part) => {
      const unquoted =
        part.length >= 2 &&
        ((part.startsWith('"') && part.endsWith('"')) ||
          (part.startsWith("'") && part.endsWith("'")))
          ? part.slice(1, -1)
          : part;
      return /^[a-zA-Z][a-zA-Z0-9-]*$/.test(unquoted) ? unquoted : cssString(unquoted);
    })
    .join(', ');

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
  if (fmt.fontFamily) parts.push(`font-family:${safeFontFamily(fmt.fontFamily)}`);
  if (fmt.fontSize) parts.push(`font-size:${fmt.fontSize}px`);
  return parts.join(';');
};

/** Render the range as an HTML `<table>` with inline styles. */
export function encodeHtml(
  state: State,
  range: Range,
  locale?: string,
  formatResolver?: ClipboardFormatResolver,
): string {
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
      const fmt = formatResolver
        ? formatResolver({ sheet: range.sheet, row: r, col: c })
        : state.format.formats.get(key);
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
              ? formatNumber(cell.value.value, fmt.numFmt, locale, state.ui.date1904 === true)
              : formatCell(cell.value, locale);
      const style = styleOf(fmt ?? undefined);
      const styleAttr = style ? ` style="${escapeHtml(style)}"` : '';
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
