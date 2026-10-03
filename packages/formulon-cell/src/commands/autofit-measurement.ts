/** Content-fit column widths and row heights. Every autofit entry point — header
 *  double-click, the structure commands, the ribbon AutoFit items and the
 *  public `autofitColWidth` / `autofitRowHeight` helpers — measures here.
 *  Results are unclamped; the store's size mutators own the min/max bounds. */
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import { normalizeFormatLocale } from '../format/locale.js';
import { FILTER_BTN_INSET, FILTER_BTN_SIZE } from '../render/geometry.js';
import {
  TABLE_HEADER_CHEVRON_INSET,
  TABLE_HEADER_CHEVRON_SIZE,
} from '../render/painters/markers.js';
import { type CellFontTheme, cellFont, cellFontCss } from '../render/painters/text.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { DEFAULT_CELL_FONT_THEME, resolveTheme } from '../theme/resolve.js';
import { formatNumber } from './format.js';
import { isHeaderRow, tableForCell } from './format-as-table.js';

/** Width the autofilter dropdown button occupies at a header's trailing edge. */
export const FILTER_DROPDOWN_RESERVED_WIDTH = FILTER_BTN_INSET + FILTER_BTN_SIZE;

/** Width the table header dropdown chevron occupies at a cell's trailing edge. */
export const TABLE_HEADER_RESERVED_WIDTH = TABLE_HEADER_CHEVRON_INSET + TABLE_HEADER_CHEVRON_SIZE;

const COL_PADDING = 16;
const MIN_COL_WIDTH = 48;

/** The cell-format fields autofit measurement reads. */
export type AutofitCellFormat = Pick<
  CellFormat,
  'fontSize' | 'fontFamily' | 'fontVertAlign' | 'bold' | 'italic' | 'numFmt' | 'wrap'
>;

export interface AutofitOptions {
  /** Inclusive cross-axis bounds to scan: rows for a column fit, columns for a
   *  row fit. The whole column / row when omitted. */
  span?: { from: number; to: number };
  /** Number-format locale (`ja`, `en-US`, ...); `en-US` when omitted. */
  locale?: string;
  /** Default cell font from the resolved theme; the theme defaults when omitted. */
  theme?: CellFontTheme;
}

export function createAutofitMeasureContext(): CanvasRenderingContext2D | null {
  const doc = globalThis.document;
  const canvas = doc?.createElement?.('canvas');
  return canvas?.getContext?.('2d') ?? null;
}

export function computeAutofitColWidth(
  state: State,
  col: number,
  ctx: CanvasRenderingContext2D | null,
  opts: AutofitOptions = {},
): number {
  const sheet = state.data.sheetIndex;
  const locale = normalizeFormatLocale(opts.locale ?? '');
  const theme = opts.theme ?? DEFAULT_CELL_FONT_THEME;
  let max = 0;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.col !== col) continue;
    if (!inSpan(parsed.row, opts.span)) continue;
    const text = autofitDisplayText(state, key, cell, locale);
    if (!text) continue;
    const font = cellFont(state.format.formats.get(key), theme, showsFormula(state, cell));
    if (ctx) ctx.font = cellFontCss(font);
    const width =
      maxExplicitLineWidth(text, ctx, font.size) +
      headerButtonWidth(state, parsed.sheet, parsed.row, parsed.col);
    if (width > max) max = width;
  }

  return Math.max(MIN_COL_WIDTH, Math.ceil(max) + COL_PADDING);
}

export function computeAutofitRowHeight(
  state: State,
  row: number,
  ctx: CanvasRenderingContext2D | null,
  opts: AutofitOptions = {},
): number {
  const sheet = state.data.sheetIndex;
  const locale = normalizeFormatLocale(opts.locale ?? '');
  const theme = opts.theme ?? DEFAULT_CELL_FONT_THEME;
  let max = state.layout.defaultRowHeight;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.row !== row) continue;
    if (!inSpan(parsed.col, opts.span)) continue;
    const text = autofitDisplayText(state, key, cell, locale);
    if (!text) continue;
    const fmt = state.format.formats.get(key);
    const font = cellFont(fmt, theme, showsFormula(state, cell));
    if (ctx) ctx.font = cellFontCss(font);
    const lineHeight = Math.round(font.size * 1.28);
    const colW = state.layout.colWidths.get(parsed.col) ?? state.layout.defaultColWidth;
    const lines = autofitLineCount(text, fmt?.wrap === true, colW, ctx, font.size);
    const height = Math.ceil(lines * lineHeight + 8);
    if (height > max) max = height;
  }

  return max;
}

/** Fitted width of `col`, measuring only rows `r0..r1` in the host's theme font. */
export function autofitColWidth(
  instance: { readonly store: SpreadsheetStore; readonly host: HTMLElement },
  col: number,
  r0: number,
  r1: number,
  locale: string,
): number {
  return computeAutofitColWidth(instance.store.getState(), col, createAutofitMeasureContext(), {
    span: { from: r0, to: r1 },
    locale,
    theme: resolveTheme(instance.host),
  });
}

/** Fitted height of `row`, measuring only columns `c0..c1` in the host's theme font. */
export function autofitRowHeight(
  instance: { readonly store: SpreadsheetStore; readonly host: HTMLElement },
  row: number,
  c0: number,
  c1: number,
  locale: string,
): number {
  return computeAutofitRowHeight(instance.store.getState(), row, createAutofitMeasureContext(), {
    span: { from: c0, to: c1 },
    locale,
    theme: resolveTheme(instance.host),
  });
}

function inSpan(index: number, span: AutofitOptions['span']): boolean {
  return !span || (index >= span.from && index <= span.to);
}

function parseCellKey(key: string): { sheet: number; row: number; col: number } | null {
  const parts = key.split(':');
  if (parts.length !== 3) return null;
  const sheet = Number(parts[0]);
  const row = Number(parts[1]);
  const col = Number(parts[2]);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
}

function autofitDisplayText(
  state: State,
  key: string,
  cell: { value: CellValue; formula: string | null },
  locale: string,
): string {
  if (showsFormula(state, cell)) return cell.formula as string;
  const fmt = state.format.formats.get(key);
  if (cell.value.kind === 'number' && fmt?.numFmt)
    return formatNumber(cell.value.value, fmt.numFmt, locale);
  return formatCell(cell.value, locale);
}

function showsFormula(state: State, cell: { formula: string | null }): boolean {
  return state.ui.showFormulas && !!cell.formula;
}

// One trailing allowance per header cell: the wider of the filter and table buttons.
function headerButtonWidth(state: State, sheet: number, row: number, col: number): number {
  const fr = state.ui.filterRange;
  const filter =
    fr && fr.sheet === sheet && fr.r0 === row && col >= fr.c0 && col <= fr.c1
      ? FILTER_DROPDOWN_RESERVED_WIDTH
      : 0;
  const table = tableForCell(state.tables.tables, sheet, row, col);
  const chevron = table && isHeaderRow(table, row, col) ? TABLE_HEADER_RESERVED_WIDTH : 0;
  return Math.max(filter, chevron);
}

function maxExplicitLineWidth(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  let max = 0;
  for (const line of text.split(/\r\n|\r|\n/)) {
    const width = measureAutofitText(line, ctx, fontSize);
    if (width > max) max = width;
  }
  return max;
}

function autofitLineCount(
  text: string,
  wrap: boolean,
  colWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const paragraphs = text.split(/\r\n|\r|\n/);
  if (!wrap) return Math.max(1, paragraphs.length);

  const available = Math.max(1, colWidth - 12);
  let lines = 0;
  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines += 1;
      continue;
    }
    lines += wrapAutofitParagraph(paragraph, available, ctx, fontSize);
  }
  return Math.max(1, lines);
}

function wrapAutofitParagraph(
  paragraph: string,
  maxWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const words = paragraph.split(/(\s+)/);
  let line = '';
  let count = 0;
  for (const word of words) {
    const candidate = line + word;
    if (measureAutofitText(candidate, ctx, fontSize) <= maxWidth || line === '') {
      line = candidate;
    } else {
      count += 1;
      line = word.trimStart();
    }
  }
  return count + (line ? 1 : 0);
}

// Character-width estimate when no canvas is available (headless test DOMs).
function measureAutofitText(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const measured = ctx ? ctx.measureText(text).width : 0;
  return measured > 0 ? measured : text.length * fontSize * 0.54;
}
