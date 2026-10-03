/** Content-fit column widths and row heights. Every autofit entry point — header
 *  double-click, the structure commands, the ribbon AutoFit items and the
 *  public `autofitColWidth` / `autofitRowHeight` helpers — measures here.
 *  Results are unclamped; the store's size mutators own the min/max bounds. */
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import { normalizeFormatLocale } from '../format/locale.js';
import { FILTER_BTN_INSET, FILTER_BTN_SIZE } from '../render/geometry.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { formatNumber } from './format.js';

/** Width the autofilter dropdown button occupies at a header's trailing edge. */
export const FILTER_DROPDOWN_RESERVED_WIDTH = FILTER_BTN_INSET + FILTER_BTN_SIZE;

const COL_PADDING = 16;
const MIN_COL_WIDTH = 48;
const DEFAULT_FONT_SIZE = 13;

/** The cell-format fields autofit measurement reads. */
export type AutofitCellFormat = Pick<
  CellFormat,
  'fontSize' | 'fontFamily' | 'bold' | 'italic' | 'numFmt' | 'wrap'
>;

export interface AutofitOptions {
  /** Inclusive cross-axis bounds to scan: rows for a column fit, columns for a
   *  row fit. The whole column / row when omitted. */
  span?: { from: number; to: number };
  /** Number-format locale (`ja`, `en-US`, ...); `en-US` when omitted. */
  locale?: string;
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
  let max = 0;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.col !== col) continue;
    if (!inSpan(parsed.row, opts.span)) continue;
    const text = autofitDisplayText(state, key, cell, locale);
    if (!text) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? DEFAULT_FONT_SIZE;
    if (ctx) ctx.font = autofitFont(fmt);
    const width =
      maxExplicitLineWidth(text, ctx, fontSize) +
      (isFilterHeaderCell(state, parsed.sheet, parsed.row, parsed.col)
        ? FILTER_DROPDOWN_RESERVED_WIDTH
        : 0);
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
  let max = state.layout.defaultRowHeight;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.row !== row) continue;
    if (!inSpan(parsed.col, opts.span)) continue;
    const text = autofitDisplayText(state, key, cell, locale);
    if (!text) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? DEFAULT_FONT_SIZE;
    if (ctx) ctx.font = autofitFont(fmt);
    const lineHeight = Math.round(fontSize * 1.28);
    const colW = state.layout.colWidths.get(parsed.col) ?? state.layout.defaultColWidth;
    const lines = autofitLineCount(text, fmt?.wrap === true, colW, ctx, fontSize);
    const height = Math.ceil(lines * lineHeight + 8);
    if (height > max) max = height;
  }

  return max;
}

/** Fitted width of `col`, measuring only rows `r0..r1`. */
export function autofitColWidth(
  instance: { readonly store: SpreadsheetStore },
  col: number,
  r0: number,
  r1: number,
  locale: string,
): number {
  return computeAutofitColWidth(instance.store.getState(), col, createAutofitMeasureContext(), {
    span: { from: r0, to: r1 },
    locale,
  });
}

/** Fitted height of `row`, measuring only columns `c0..c1`. */
export function autofitRowHeight(
  instance: { readonly store: SpreadsheetStore },
  row: number,
  c0: number,
  c1: number,
  locale: string,
): number {
  return computeAutofitRowHeight(instance.store.getState(), row, createAutofitMeasureContext(), {
    span: { from: c0, to: c1 },
    locale,
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
  if (state.ui.showFormulas && cell.formula) return cell.formula;
  const fmt = state.format.formats.get(key);
  if (cell.value.kind === 'number' && fmt?.numFmt)
    return formatNumber(cell.value.value, fmt.numFmt, locale);
  return formatCell(cell.value, locale);
}

function isFilterHeaderCell(state: State, sheet: number, row: number, col: number): boolean {
  const fr = state.ui.filterRange;
  return !!fr && fr.sheet === sheet && fr.r0 === row && col >= fr.c0 && col <= fr.c1;
}

function autofitFont(format: AutofitCellFormat | undefined): string {
  const styleSlant = format?.italic ? 'italic ' : '';
  const weight = format?.bold ? 700 : 400;
  const size = format?.fontSize ?? DEFAULT_FONT_SIZE;
  const family = format?.fontFamily ?? 'system-ui, sans-serif';
  return `${styleSlant}${weight} ${size}px ${fontFamilyCss(family)}`;
}

function fontFamilyCss(family: string): string {
  return family
    .split(',')
    .map((part) => {
      const trimmed = part.trim();
      if (/^["'].*["']$/.test(trimmed) || /^[a-z-]+$/i.test(trimmed)) return trimmed;
      return `"${trimmed.replace(/"/g, '\\"')}"`;
    })
    .join(', ');
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
