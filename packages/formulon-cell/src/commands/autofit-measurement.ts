import { parseAddrKey } from '../engine/address.js';
import { effectiveCellFormat } from '../engine/cell-format-projection.js';
import { fontPointsToCssPx } from '../engine/font-units.js';
import { macDefaultRowPt, macWrapLinePt } from '../engine/mac-row-model.js';
import { phoneticGuideSize } from '../engine/phonetic-sync.js';
import { pivotStyleLayer } from '../engine/pivot-style.js';
/** Content-fit column widths and row heights. Every autofit entry point — header
 *  double-click, the structure commands, the ribbon AutoFit items and the
 *  public `autofitColWidth` / `autofitRowHeight` helpers — measures here.
 *  Results are unclamped; the store's size mutators own the min/max bounds. */
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { normalizeFormatLocale } from '../format/locale.js';
import { styleElementFormatSource } from '../render/grid/table-format.js';
import { FILTER_BUTTON_SIZE, MAC_PIVOT_BUTTON_SIZE } from '../render/painters/filter-button.js';
import {
  TABLE_HEADER_CHEVRON_INSET,
  TABLE_HEADER_CHEVRON_SIZE,
} from '../render/painters/markers.js';
import { pivotOutlineIndent } from '../render/painters/pivot-outline-button.js';
import {
  type CellFont,
  type CellFontTheme,
  cellFont,
  cellFontCss,
} from '../render/painters/text.js';
import {
  cellLinePitch,
  cellTextWidth,
  paintedFontSize,
  wholePointAdvances,
} from '../render/text-metrics.js';
import { cellTextHorizontalMetrics, wrapText } from '../render/text-wrap.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { DEFAULT_CELL_FONT_THEME, resolveTheme } from '../theme/resolve.js';
import { filterButtonHidden } from './filter-ranges.js';
import { formatNumber } from './format.js';
import { showsHeaderFilterButton, tableForCell } from './format-as-table.js';

/** Width the in-cell AutoFilter button occupies at a header cell's trailing
 *  edge: the button plus its one-pixel inset. */
export const FILTER_DROPDOWN_RESERVED_WIDTH = FILTER_BUTTON_SIZE + 1;

/** Width the table header dropdown chevron occupies at a cell's trailing edge. */
export const TABLE_HEADER_RESERVED_WIDTH = TABLE_HEADER_CHEVRON_INSET + TABLE_HEADER_CHEVRON_SIZE;

const COL_PADDING = 16;

/** The cell-format fields autofit measurement reads. */
export type AutofitCellFormat = Pick<
  CellFormat,
  | 'fontSize'
  | 'fontFamily'
  | 'fontVertAlign'
  | 'bold'
  | 'italic'
  | 'numFmt'
  | 'wrap'
  | 'indent'
  | 'rotation'
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
  workbookDisplayFont?: { name: string; size: number } | null;
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

/** Excel leaves merged cells out of AutoFit on either axis. */
const isMerged = (state: State, key: string): boolean =>
  state.merges.byCell.has(key) || state.merges.byAnchor.has(key);

/** Column AutoFit, as Excel does it: wrapped and merged cells do not count, and
 *  a column with nothing to measure takes the sheet's default width. */
export function computeAutofitColWidth(
  state: State,
  col: number,
  ctx: CanvasRenderingContext2D | null,
  opts: AutofitOptions = {},
): number {
  const sheet = state.data.sheetIndex;
  const locale = normalizeFormatLocale(opts.locale ?? '');
  const theme = opts.theme ?? DEFAULT_CELL_FONT_THEME;
  let max: number | null = null;
  const styleAt = styleElementFormatSource(state, opts.workbook ?? null);
  const pivots = pivotStyleLayer(opts.workbook ?? null, state);

  for (const [key, cell] of state.data.cells) {
    const parsed = parseAddrKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.col !== col) continue;
    if (!inSpan(parsed.row, opts.span) || isMerged(state, key)) continue;
    const format = measurementFormat(state, key, parsed, opts, cell, styleAt);
    if (format?.wrap === true) continue;
    const text = autofitDisplayText(state, cell, format, locale);
    if (!text) continue;
    const formulaDisplay = showsFormula(state, cell);
    const font = cellFont(format, theme, formulaDisplay);
    const width =
      maxExplicitLineWidth(text, textMeasure(ctx, font, state.layout)) +
      headerButtonWidth(state, parsed.sheet, parsed.row, parsed.col, pivots.buttons.has(key)) +
      pivotOutlineIndent(
        pivots.outlineAt(parsed.row, parsed.col)?.steps ?? 0,
        fontPointsToCssPx(1),
      );
    if (max === null || width > max) max = width;
  }

  return max === null ? state.layout.defaultColWidth : Math.ceil(max) + COL_PADDING;
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
  const defaultHeight = state.layout.defaultRowHeight;
  let max: number | null = null;
  let wrapped = false;
  // The default height is one Normal-font line plus the row's padding, so a
  // single line of the Normal font fits back to exactly the default height.
  const normalFont = cellFont(
    formatForMeasurement(undefined, opts.workbookDisplayFont, false),
    theme,
    false,
  );
  const padding = Math.max(0, defaultHeight - cellLinePitch(normalFont.size));
  const mac = wholePointAdvances(state.layout);
  const styleAt = styleElementFormatSource(state, opts.workbook ?? null);

  for (const [key, cell] of state.data.cells) {
    const parsed = parseAddrKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.row !== row) continue;
    if (!inSpan(parsed.col, opts.span) || isMerged(state, key)) continue;
    const format = measurementFormat(state, key, parsed, opts, cell, styleAt);
    const text = autofitDisplayText(state, cell, format, locale);
    if (!text) continue;
    const formulaDisplay = showsFormula(state, cell);
    const font = cellFont(format, theme, formulaDisplay);
    const lineHeight = cellLinePitch(font.size);
    const colW = state.layout.colWidths.get(parsed.col) ?? state.layout.defaultColWidth;
    const lines = autofitLineCount(
      text,
      format?.wrap === true,
      colW,
      textMeasure(ctx, font, state.layout),
      locale,
      format?.indent,
    );
    // A shown phonetic guide sits above the first line and needs room of its own.
    const guide =
      format?.showPhonetic && format.phonetic?.length && !formulaDisplay
        ? phoneticGuideSize(font.size)
        : 0;
    const extent = rotatedExtent(
      format?.rotation,
      text,
      lines * lineHeight + guide,
      lineHeight,
      textMeasure(ctx, font, state.layout),
    );
    const family = format?.fontFamily ?? opts.workbookDisplayFont?.name ?? '';
    const sizePt = Math.round(font.size * POINTS_PER_CSS_PX * 100) / 100;
    const tabulated =
      mac && !format?.rotation && guide === 0 ? macDefaultRowPt(family, sizePt) : null;
    // Wrapped text takes whole lines of the font's wrap pitch, one line included.
    const wrapPitch =
      tabulated !== null && format?.wrap === true ? macWrapLinePt(family, sizePt) : null;
    const height =
      tabulated === null
        ? Math.ceil(extent) + padding
        : wrapPitch !== null
          ? Math.max(1, lines) * fontPointsToCssPx(wrapPitch)
          : macLinesHeight(fontPointsToCssPx(tabulated), lines);
    if (wrapPitch !== null) wrapped = true;
    if (max === null || height > max) max = height;
  }

  // A Mac row holding wrapped text fits its cells' own lines, even below the
  // default row; any other row keeps the default height as its floor.
  if (max === null) return defaultHeight;
  return mac && wrapped ? max : Math.max(defaultHeight, max);
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
  styleAt: (row: number, col: number) => CellFormat | undefined,
): CellFormat | undefined {
  const formulaDisplay = showsFormula(state, cell);
  const direct = opts.workbook
    ? effectiveCellFormat(opts.workbook, state, addr)
    : state.format.formats.get(key);
  // A table or PivotTable style's bold measures as the cell paints it.
  const style = styleAt(addr.row, addr.col);
  const effective =
    style?.bold && direct?.bold !== true ? { ...direct, bold: true } : (direct ?? undefined);
  return formatForMeasurement(effective, opts.workbookDisplayFont, formulaDisplay);
}

function formatForMeasurement(
  format: CellFormat | undefined,
  workbookDisplayFont: AutofitOptions['workbookDisplayFont'],
  formulaDisplay: boolean,
): AutofitCellFormat {
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
function headerButtonWidth(
  state: State,
  sheet: number,
  row: number,
  col: number,
  pivotButton: boolean,
): number {
  if (pivotButton) {
    return wholePointAdvances(state.layout)
      ? fontPointsToCssPx(MAC_PIVOT_BUTTON_SIZE) + 1
      : FILTER_DROPDOWN_RESERVED_WIDTH;
  }
  const filter =
    state.ui.filterRanges.some(
      (fr) => fr.sheet === sheet && fr.r0 === row && col >= fr.c0 && col <= fr.c1,
    ) && !filterButtonHidden(state, sheet, row, col)
      ? FILTER_DROPDOWN_RESERVED_WIDTH
      : 0;
  const table = tableForCell(state.tables.tables, sheet, row, col);
  const chevron =
    table && showsHeaderFilterButton(table, row, col) ? TABLE_HEADER_RESERVED_WIDTH : 0;
  return Math.max(filter, chevron);
}

/** Points per canonical CSS px. */
const POINTS_PER_CSS_PX = 0.75;

/** Mac Excel fits one line of a tabulated font to that font's default row and
 *  gives each line of a wrapped block one point more (Arial 11: 14 pt for one
 *  line, 30 / 45 pt for two / three). */
const macLinesHeight = (rowPx: number, lines: number): number =>
  lines <= 1 ? rowPx : lines * (rowPx + fontPointsToCssPx(1));

/** Vertical extent of a cell's text block. Rotated text stands its run up in
 *  the row (Excel: 'rot' at 90 degrees takes a 14 pt row to 19 pt); stacked
 *  text puts one character per line. */
function rotatedExtent(
  rotation: CellFormat['rotation'],
  text: string,
  blockHeight: number,
  lineHeight: number,
  measure: (text: string) => number,
): number {
  if (rotation === 'vertical') return [...text.replace(/\r?\n/g, '')].length * lineHeight;
  if (!rotation) return blockHeight;
  const angle = (Math.abs(rotation) * Math.PI) / 180;
  const run = maxExplicitLineWidth(text, measure);
  return run * Math.sin(angle) + blockHeight * Math.cos(angle);
}

function maxExplicitLineWidth(text: string, measure: (text: string) => number): number {
  let max = 0;
  for (const line of text.split(/\r\n|\r|\n/)) {
    const width = measure(line);
    if (width > max) max = width;
  }
  return max;
}

function autofitLineCount(
  text: string,
  wrap: boolean,
  colWidth: number,
  measure: (text: string) => number,
  locale: string,
  indent = 0,
): number {
  const horizontal = cellTextHorizontalMetrics(colWidth, indent, 1);
  return wrapText(text, {
    maxWidth: wrap ? horizontal.availableWidth : Infinity,
    measureText: measure,
    locale,
  }).length;
}

/** Run widths in stored CSS px for `font`, laid out as the grid paints them:
 *  on the point model the advances round at the screen size, so the run is
 *  measured there and scaled back. Without a canvas (headless test DOMs) a
 *  character-width estimate stands in. */
function textMeasure(
  ctx: CanvasRenderingContext2D | null,
  font: CellFont,
  layout: State['layout'],
): (text: string) => number {
  const wholePoint = wholePointAdvances(layout);
  const scale = wholePoint ? (layout.screenScale ?? 1) : 1;
  if (ctx) ctx.font = cellFontCss(font, paintedFontSize(font.size, scale, wholePoint));
  return (text) => {
    const measured = ctx ? cellTextWidth(ctx, text, wholePoint) / scale : 0;
    return measured > 0 ? measured : text.length * font.size * 0.54;
  };
}
