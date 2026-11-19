/**
 * Desktop-spreadsheet-compatible HTML clipboard encoder. The output is a `<table>`
 * with inline styles for bold/italic/underline/strike/align/color/fill —
 * Spreadsheets parse these on paste.
 */

import { addrKey } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import { formatCell } from '../../engine/value.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { safeFontFamily } from '../../format/font-family.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import { cssString, escapeHtml } from '../../text-escape.js';
import { formatNumber } from '../format.js';
import { type History, transactIn } from '../history.js';
import { type ChangeBatchResult, refusedBatch } from '../interaction-policy.js';
import { recordFormatChange } from '../slice-history.js';
import { pasteTSV } from './paste.js';
import type { ClipboardFormatResolver } from './snapshot.js';
import { encodeTSV } from './tsv.js';

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
  if (fmt.fontFamily) parts.push(`font-family:${safeFontFamily(fmt.fontFamily, cssString)}`);
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

/** One cell of a pasted HTML table: its text and the one format Excel keeps. */
export interface HtmlTableCell {
  text: string;
  bold: boolean;
}

const boldWeight = (weight: string): boolean =>
  weight === 'bold' || weight === 'bolder' || Number.parseInt(weight, 10) >= 600;

/** Whether all of `cell`'s text is set bold, by its own style or by a
 *  `<b>` / `<strong>` / bold-weight element around it. */
const htmlCellBold = (cell: Element): boolean => {
  const text = (cell.textContent ?? '').trim();
  if (!text) return false;
  if (boldWeight((cell as HTMLElement).style?.fontWeight ?? '')) return true;
  for (const el of Array.from(cell.querySelectorAll('*'))) {
    const bold =
      el.tagName === 'B' ||
      el.tagName === 'STRONG' ||
      boldWeight((el as HTMLElement).style?.fontWeight ?? '');
    if (bold && (el.textContent ?? '').trim() === text) return true;
  }
  return false;
};

/**
 * Read the first `<table>` of an HTML clipboard payload as Excel does: each
 * row's cells left to right in reading order, `colspan` and `rowspan` ignored
 * (a spanned cell takes one cell and opens no merge), text and bold kept,
 * fills and colours dropped. Null when the payload holds no table or no
 * `DOMParser` is available.
 */
export function parseHtmlTable(html: string): HtmlTableCell[][] | null {
  const Parser = (globalThis as { DOMParser?: typeof DOMParser }).DOMParser;
  if (!Parser) return null;
  const table = new Parser().parseFromString(html, 'text/html').querySelector('table');
  if (!table) return null;
  return Array.from(table.querySelectorAll('tr')).map((tr) =>
    Array.from(tr.children)
      .filter((cell) => cell.tagName === 'TD' || cell.tagName === 'TH')
      .map((cell) => ({ text: (cell.textContent ?? '').trim(), bold: htmlCellBold(cell) })),
  );
}

/**
 * Paste an HTML table at the active cell: the cells as {@link parseHtmlTable}
 * reads them, typed as pasted text is, with bold applied where the source
 * set it. Null when the payload holds no table.
 */
export function pasteHtml(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  html: string,
  history: History | null = null,
): ChangeBatchResult {
  const rows = parseHtmlTable(html);
  if (!rows || rows.length === 0)
    return refusedBatch({
      code: 'invalid',
      reason: 'nothing to paste',
      reasonKey: 'clipboardEmpty',
    });
  return transactIn(history, () => {
    const result = pasteTSV(
      state,
      store,
      wb,
      encodeTSV(rows.map((row) => row.map((cell) => cell.text))),
      history,
    );
    if (result.status !== 'applied' || !result.range) return result;
    const { sheet, r0, c0 } = result.range;
    const bold = rows.flatMap((row, r) =>
      row.flatMap((cell, c) => (cell.bold ? [{ sheet, row: r0 + r, col: c0 + c }] : [])),
    );
    if (bold.length > 0) {
      recordFormatChange(history, store, () => {
        for (const addr of bold) mutators.setCellFormat(store, addr, { bold: true });
      });
    }
    return result;
  });
}

/** Marks the HTML this workbook writes to the system clipboard. That HTML
 *  shows numbers as displayed, so a paste of it goes through the lossless
 *  plain text written beside it instead of {@link pasteHtml}. */
const CLIPBOARD_HTML_GENERATOR = '<meta name="generator" content="formulon-cell">';

/** The clipboard HTML for `tables`, marked as this workbook's own. */
export const clipboardHtml = (tables: string): string => `${CLIPBOARD_HTML_GENERATOR}${tables}`;

/** Whether `html` is clipboard HTML this workbook wrote. */
export const isOwnClipboardHtml = (html: string): boolean =>
  html.includes(CLIPBOARD_HTML_GENERATOR);
