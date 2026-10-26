import { parseAddrKey } from '../engine/address.js';
import { effectiveCellFormatWithPending } from '../engine/cell-format-sync.js';
import { fontPointsToCssPx } from '../engine/font-units.js';
import { phoneticGuideSize } from '../engine/phonetic-sync.js';
/** Content-fit column widths and row heights. Every autofit entry point — header
 *  double-click, the structure commands, the ribbon AutoFit items and the
 *  public `autofitColWidth` / `autofitRowHeight` helpers — measures here.
 *  Results are unclamped; the store's size mutators own the min/max bounds. */
import type { CellValue, FontRecord } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { normalizeFormatLocale } from '../format/locale.js';
import { FILTER_BTN_INSET, FILTER_BTN_SIZE } from '../render/geometry.js';
import {
  TABLE_HEADER_CHEVRON_INSET,
  TABLE_HEADER_CHEVRON_SIZE,
} from '../render/painters/markers.js';
import { type CellFontTheme, cellFont, cellFontCss } from '../render/painters/text.js';
import { cellTextHorizontalMetrics, wrapText } from '../render/text-wrap.js';
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
  'fontSize' | 'fontFamily' | 'fontVertAlign' | 'bold' | 'italic' | 'numFmt' | 'wrap' | 'indent'
>;

export interface AutofitOptions {
  /** Inclusive cross-axis bounds to scan: rows for a column fit, columns for a
   *  row fit. The whole column / row when omitted. */
  span?: { from: number; to: number };
  /** Number-format locale (`ja`, `en-US`, ...); `en-US` when omitted. */
  locale?: string;
  /** Default cell font from the resolved theme; the theme defaults when omitted. */
  theme?: CellFontTheme;
  /** Workbook Normal/display font in native points. When supplied, missing
   *  cell family/size fields are measured with this font instead of the UI
   *  theme fallback. */
  workbookDisplayFont?: Pick<FontRecord, 'name' | 'size'> | null;
  /** Workbook used to resolve row/column/whole-sheet inherited formatting.
   *  The resolved projection is read-only and is never written into the
   *  store's direct format map. */
  workbook?: WorkbookHandle | null;
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
    const parsed = parseAddrKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.col !== col) continue;
    if (!inSpan(parsed.row, opts.span)) continue;
    const format = measurementFormat(state, key, parsed, opts, cell);
    const text = autofitDisplayText(state, cell, format, locale);
    if (!text) continue;
    const formulaDisplay = showsFormula(state, cell);
    const font = cellFont(format, theme, formulaDisplay);
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
    const parsed = parseAddrKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.row !== row) continue;
    if (!inSpan(parsed.col, opts.span)) continue;
    const format = measurementFormat(state, key, parsed, opts, cell);
    const text = autofitDisplayText(state, cell, format, locale);
    if (!text) continue;
    const formulaDisplay = showsFormula(state, cell);
    const font = cellFont(format, theme, formulaDisplay);
    if (ctx) ctx.font = cellFontCss(font);
    const lineHeight = Math.round(font.size * 1.28);
    const colW = state.layout.colWidths.get(parsed.col) ?? state.layout.defaultColWidth;
    const lines = autofitLineCount(
      text,
      format?.wrap === true,
      colW,
      ctx,
      font.size,
      locale,
      format?.indent,
    );
    // A shown phonetic guide sits above the first line and needs room of its own.
    const guide =
      format?.showPhonetic && format.phonetic?.length && !formulaDisplay
        ? phoneticGuideSize(font.size)
        : 0;
    const height = Math.ceil(lines * lineHeight + 8 + guide);
    if (height > max) max = height;
  }

  return max;
}

/** Fitted width of `col`, measuring only rows `r0..r1` in the host's theme font. */
export function autofitColWidth(
  instance: {
    readonly store: SpreadsheetStore;
    readonly host: HTMLElement;
    readonly workbook?: { readonly workbookDisplayFont: AutofitOptions['workbookDisplayFont'] };
    readonly workbookHandle?: WorkbookHandle;
  },
  col: number,
  r0: number,
  r1: number,
  locale: string,
): number {
  return computeAutofitColWidth(instance.store.getState(), col, createAutofitMeasureContext(), {
    span: { from: r0, to: r1 },
    locale,
    theme: resolveTheme(instance.host),
    workbookDisplayFont: instance.workbook?.workbookDisplayFont,
    workbook: instance.workbookHandle ?? workbookHandleFromInstance(instance.workbook),
  });
}

/** Fitted height of `row`, measuring only columns `c0..c1` in the host's theme font. */
export function autofitRowHeight(
  instance: {
    readonly store: SpreadsheetStore;
    readonly host: HTMLElement;
    readonly workbook?: { readonly workbookDisplayFont: AutofitOptions['workbookDisplayFont'] };
    readonly workbookHandle?: WorkbookHandle;
  },
  row: number,
  c0: number,
  c1: number,
  locale: string,
): number {
  return computeAutofitRowHeight(instance.store.getState(), row, createAutofitMeasureContext(), {
    span: { from: c0, to: c1 },
    locale,
    theme: resolveTheme(instance.host),
    workbookDisplayFont: instance.workbook?.workbookDisplayFont,
    workbook: instance.workbookHandle ?? workbookHandleFromInstance(instance.workbook),
  });
}

function workbookHandleFromInstance(
  workbook:
    | ({ readonly workbookDisplayFont: AutofitOptions['workbookDisplayFont'] } & {
        readonly getEffectiveStyle?: unknown;
      })
    | undefined,
): WorkbookHandle | undefined {
  return workbook && typeof workbook.getEffectiveStyle === 'function'
    ? (workbook as unknown as WorkbookHandle)
    : undefined;
}

function inSpan(index: number, span: AutofitOptions['span']): boolean {
  return !span || (index >= span.from && index <= span.to);
}

function autofitDisplayText(
  state: State,
  cell: { value: CellValue; formula: string | null },
  format: CellFormat | undefined,
  locale: string,
): string {
  if (showsFormula(state, cell)) return cell.formula as string;
  if (cell.value.kind === 'number' && format?.numFmt)
    return formatNumber(cell.value.value, format.numFmt, locale, state.ui.date1904 === true);
  return formatCell(cell.value, locale);
}

function showsFormula(state: State, cell: { formula: string | null }): boolean {
  return state.ui.showFormulas && !!cell.formula;
}

function measurementFormat(
  state: State,
  key: string,
  addr: { sheet: number; row: number; col: number },
  opts: AutofitOptions,
  cell: { formula: string | null },
): CellFormat | undefined {
  const formulaDisplay = showsFormula(state, cell);
  const effective = opts.workbook
    ? effectiveCellFormatWithPending(opts.workbook, state, addr)
    : state.format.formats.get(key);
  return formatForMeasurement(effective ?? undefined, opts.workbookDisplayFont, formulaDisplay);
}

function formatForMeasurement(
  format: CellFormat | undefined,
  workbookDisplayFont: AutofitOptions['workbookDisplayFont'],
  formulaDisplay: boolean,
): Pick<
  CellFormat,
  'fontFamily' | 'fontSize' | 'fontVertAlign' | 'bold' | 'italic' | 'numFmt' | 'wrap' | 'indent'
> {
  if (!workbookDisplayFont) return format ?? {};
  return {
    ...(format ?? {}),
    ...(format?.fontSize === undefined
      ? { fontSize: fontPointsToCssPx(workbookDisplayFont.size) }
      : {}),
    // Formula display intentionally stays on the monospace theme. For cell
    // values, the workbook Normal family is the effective fallback.
    ...(formulaDisplay || format?.fontFamily !== undefined
      ? {}
      : { fontFamily: workbookDisplayFont.name }),
  };
}

// One trailing allowance per header cell: the wider of the filter and table buttons.
function headerButtonWidth(state: State, sheet: number, row: number, col: number): number {
  const filter = state.ui.filterRanges.some(
    (fr) => fr.sheet === sheet && fr.r0 === row && col >= fr.c0 && col <= fr.c1,
  )
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
  locale: string,
  indent = 0,
): number {
  const horizontal = cellTextHorizontalMetrics(colWidth, indent, 1);
  return wrapText(text, {
    maxWidth: wrap ? horizontal.availableWidth : Infinity,
    measureText: (value) => measureAutofitText(value, ctx, fontSize),
    locale,
  }).length;
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
