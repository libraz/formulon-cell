import { DEFAULT_THEME_ACCENTS } from '../charts/series-colors.js';
import {
  type CustomTableStyle,
  customPivotTableStyleById,
  pivotTableStyleAssignment,
} from '../commands/format-as-table.js';
import type { CellFormat, TablesSlice } from '../store/types.js';
import { addrKey } from './address.js';
import { PIVOT_CELL_KIND } from './pivot-cell-kind.js';
import {
  composeStyleElements,
  PIVOT_ELEMENT_ORDER,
  type StyleElementFormats,
  type StyleElementRegion,
} from './table-style-elements.js';
import { customStyleElementFormats } from './table-style-sync.js';
import { cssColorFromArgb, tintCssColor } from './theme-color.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { PivotStyleDefinition } from './workbook-handle-pivot.js';

/**
 * PivotTable styles resolved for painting. A report's cells keep their own
 * (Normal) formats; its look comes from the style named on the report, as a
 * set of style elements over the report's regions: header rows, grand total
 * row, outer-item subheading rows, subtotal rows, stripes and the first
 * column. Regions are read off the engine's projected layout cell kinds.
 *
 * Built-in styles come in families of seven sharing one element shape, each
 * member tinted from one theme colour: dk1, then accent1..accent6
 * (PivotStyleLight16 is the accent1 member of Light15..Light21).
 */

/** What the renderer needs about the PivotTables of one sheet. */
export interface PivotStyleLayer {
  /** Each report's rectangle; gridlines inside one are not drawn. */
  readonly ranges: readonly Range[];
  /** Field caption cells ("Row Labels", "Column Labels") carrying a drop-down button. */
  readonly buttons: ReadonlySet<string>;
  /** The style format of a report cell; undefined outside every report. */
  formatAt(row: number, col: number): CellFormat | undefined;
  /** Outline placement of a compact row label cell; undefined elsewhere. */
  outlineAt(row: number, col: number): PivotRowOutline | undefined;
  /** The projected cells of every report that has fields, as last built. */
  readonly reports: readonly PivotReportCells[];
  /** Reports with no field laid out yet, drawn as a placeholder over the
   *  area they reserve. */
  readonly placeholders: readonly PivotPlaceholder[];
}

/** A report's rectangle (report filter block included) and its non-blank cells. */
export interface PivotReportCells {
  readonly pivotIndex: number;
  readonly range: Range;
  readonly cells: readonly PivotLayoutCell[];
}

/** An empty report reserves this many columns and rows from its anchor. */
export const PIVOT_PLACEHOLDER_COLS = 3;
export const PIVOT_PLACEHOLDER_ROWS = 18;

/** An empty report's reserved area and name. */
export interface PivotPlaceholder {
  readonly pivotIndex: number;
  readonly range: Range;
  readonly name: string;
}

/** Where a compact row label sits in the outline. */
export interface PivotRowOutline {
  /** Outline steps the label text is pushed in by: its depth, plus one when
   *  the report shows collapse buttons. */
  steps: number;
  /** The item heads a group of deeper items and carries a collapse button. */
  collapse: boolean;
}

const EMPTY_LAYER: PivotStyleLayer = {
  ranges: [],
  buttons: new Set(),
  formatAt: () => undefined,
  outlineAt: () => undefined,
  placeholders: [],
  reports: [],
};

type PivotStyleReader = Pick<
  WorkbookHandle,
  | 'capabilities'
  | 'pivotBounds'
  | 'pivotCells'
  | 'getPivotStyleInfo'
  | 'getTheme'
  | 'pivotTableName'
>;

const BUILTIN_PIVOT_STYLE = /^PivotStyle(Light|Medium|Dark)(\d+)$/;

/** The elements of a built-in PivotTable style; null for any other name. */
export function builtinPivotStyleElements(
  name: string,
  themeColors?: readonly number[],
): StyleElementFormats | null {
  const match = name.match(BUILTIN_PIVOT_STYLE);
  if (!match) return null;
  const index = Number(match[2]);
  if (!Number.isInteger(index) || index < 1) return null;
  const slot = (index - 1) % 7;
  const base =
    themeColors && themeColors.length >= 10
      ? cssColorFromArgb(themeColors[slot === 0 ? 0 : 3 + slot] ?? 0)
      : slot === 0
        ? '#000000'
        : (DEFAULT_THEME_ACCENTS[slot - 1] ?? '#000000');
  const band = tintCssColor(base, 0.8) ?? base;
  const rule = { style: 'thin' as const, color: tintCssColor(base, 0.4) ?? base };
  return {
    headerRow: { fill: band, bold: true, borders: { bottom: rule } },
    totalRow: { fill: band, bold: true, borders: { top: rule } },
    firstRowSubheading: { bold: true, borders: { top: rule } },
  };
}

/** One projected layout cell, as far as the style regions need it. */
export interface PivotLayoutCell {
  row: number;
  col: number;
  kind: number;
  /** Header nesting depth: 0 for the outermost row field's items. */
  depth: number;
  /** Source or data field of the cell; empty for captions. */
  fieldName?: string;
  /** The cell's text as shown. */
  text?: string;
}

/** Report filter rows above a report: each holds a field name and its
 *  selection side by side; a blank row then separates them from the report. */
export function pivotPageRows(range: Range, cells: readonly PivotLayoutCell[]): number {
  const at = new Map(cells.map((c) => [`${c.row}:${c.col}`, c] as const));
  let rows = 0;
  for (let row = range.r0; row <= range.r1; row += 1) {
    const name = at.get(`${row}:${range.c0}`);
    const value = at.get(`${row}:${range.c0 + 1}`);
    if (
      name?.kind !== PIVOT_CELL_KIND.Header ||
      !name.fieldName ||
      value?.kind !== PIVOT_CELL_KIND.Header ||
      value.fieldName !== name.fieldName
    ) {
      break;
    }
    rows += 1;
  }
  return rows;
}

/** First row of the report proper, below any report filter block. */
export const pivotReportTop = (range: Range, cells: readonly PivotLayoutCell[]): number => {
  const pages = pivotPageRows(range, cells);
  return range.r0 + (pages > 0 ? pages + 1 : 0);
};

const SUBHEADINGS = ['firstRowSubheading', 'secondRowSubheading', 'thirdRowSubheading'] as const;
const SUBTOTALS = ['firstSubtotalRow', 'secondSubtotalRow', 'thirdSubtotalRow'] as const;

/** The leftmost cell of each row. */
function leadCells(cells: readonly PivotLayoutCell[]): Map<number, PivotLayoutCell> {
  const lead = new Map<number, PivotLayoutCell>();
  for (const cell of cells) {
    const seen = lead.get(cell.row);
    if (!seen || cell.col < seen.col) lead.set(cell.row, cell);
  }
  return lead;
}

/** Depth of the innermost row field's items. */
const leafDepthOf = (cells: readonly PivotLayoutCell[]): number =>
  Math.max(0, ...cells.filter((c) => c.kind === PIVOT_CELL_KIND.RowLabel).map((c) => c.depth));

/** A subtotal row followed by deeper items heads its group (subtotals at the
 *  top); otherwise it closes the group. */
const closesGroup = (cell: PivotLayoutCell, next: PivotLayoutCell | undefined): boolean =>
  cell.kind === PIVOT_CELL_KIND.RowSubtotal && (!next || next.depth <= cell.depth);

/** Outline placement of the row labels in a report's first column. */
export function pivotRowOutline(
  range: Range,
  cells: readonly PivotLayoutCell[],
): { row: number; col: number; outline: PivotRowOutline }[] {
  const lead = leadCells(cells);
  const leafDepth = leafDepthOf(cells);
  const labels = cells.filter(
    (c) =>
      c.col === range.c0 &&
      (c.kind === PIVOT_CELL_KIND.RowLabel || c.kind === PIVOT_CELL_KIND.RowSubtotal),
  );
  const heads = (c: PivotLayoutCell): boolean =>
    c.depth < leafDepth && !closesGroup(c, lead.get(c.row + 1));
  const gutter = labels.some(heads) ? 1 : 0;
  return labels.map((c) => ({
    row: c.row,
    col: c.col,
    outline: { steps: c.depth + gutter, collapse: heads(c) },
  }));
}

const isRowLabelKind = (kind: number): boolean =>
  kind === PIVOT_CELL_KIND.RowLabel ||
  kind === PIVOT_CELL_KIND.RowSubtotal ||
  kind === PIVOT_CELL_KIND.GrandTotal;

/** Element regions and caption-button cells of one report, from its
 *  rectangle and its non-blank projected cells. */
export function pivotStyleRegions(
  range: Range,
  cells: readonly PivotLayoutCell[],
  info: Pick<PivotStyleDefinition, 'showRowHeaders' | 'showColHeaders' | 'showRowStripes'>,
): { regions: StyleElementRegion[]; buttons: { row: number; col: number }[] } {
  const { r1, c0, c1 } = range;
  const pages = pivotPageRows(range, cells);
  const r0 = pivotReportTop(range, cells);
  const lead = leadCells(cells.filter((c) => c.row >= r0));
  let bodyTop = r1 + 1;
  for (let row = r0; row <= r1; row += 1) {
    const kind = lead.get(row)?.kind;
    if (kind !== undefined && (isRowLabelKind(kind) || kind === PIVOT_CELL_KIND.Data)) {
      bodyTop = row;
      break;
    }
  }
  const leafDepth = leafDepthOf(cells);
  const full = (type: StyleElementRegion['type'], row: number): StyleElementRegion => ({
    type,
    r0: row,
    r1: row,
    c0,
    c1,
  });
  const regions: StyleElementRegion[] = [{ type: 'wholeTable', r0, r1, c0, c1 }];
  if (info.showColHeaders && bodyTop > r0) {
    regions.push({ type: 'headerRow', r0, r1: bodyTop - 1, c0, c1 });
  }
  let stripe = 0;
  for (let row = bodyTop; row <= r1; row += 1) {
    const cell = lead.get(row);
    if (cell?.kind === PIVOT_CELL_KIND.GrandTotal) {
      regions.push(full('totalRow', row));
      continue;
    }
    if (info.showRowStripes) {
      regions.push(full(stripe % 2 === 0 ? 'firstRowStripe' : 'secondRowStripe', row));
      stripe += 1;
    }
    if (!cell || cell.depth >= leafDepth) continue;
    const level = Math.min(cell.depth, 2);
    if (closesGroup(cell, lead.get(row + 1)))
      regions.push(full(SUBTOTALS[level] ?? 'thirdSubtotalRow', row));
    else if (info.showRowHeaders) {
      regions.push(full(SUBHEADINGS[level] ?? 'thirdRowSubheading', row));
    }
  }
  if (info.showRowHeaders && bodyTop <= r1) {
    regions.push({ type: 'firstColumn', r0: bodyTop, r1, c0, c1: c0 });
  }
  const isCaption = (row: number, col: number): boolean =>
    cells.some((c) => c.row === row && c.col === col && c.kind === PIVOT_CELL_KIND.Header);
  const buttons: { row: number; col: number }[] = [];
  for (let i = 0; i < pages; i += 1) buttons.push({ row: range.r0 + i, col: c0 + 1 });
  if (cells.some((c) => c.kind === PIVOT_CELL_KIND.RowLabel) && isCaption(bodyTop - 1, c0)) {
    buttons.push({ row: bodyTop - 1, col: c0 });
  }
  let firstColLabel: PivotLayoutCell | undefined;
  for (const c of cells) {
    if (c.kind !== PIVOT_CELL_KIND.ColLabel) continue;
    if (
      !firstColLabel ||
      c.row < firstColLabel.row ||
      (c.row === firstColLabel.row && c.col < firstColLabel.col)
    ) {
      firstColLabel = c;
    }
  }
  if (firstColLabel && isCaption(firstColLabel.row - 1, firstColLabel.col)) {
    buttons.push({ row: firstColLabel.row - 1, col: firstColLabel.col });
  }
  return { regions, buttons };
}

/** The elements of the style a report carries: a custom style assigned in the
 *  store or named on the report, else a built-in one. */
function reportStyleElements(
  tables: Pick<TablesSlice, 'customPivotTableStyles' | 'pivotTableStyles'>,
  sheet: number,
  pivotIndex: number,
  name: string,
  themeColors: readonly number[] | undefined,
): StyleElementFormats | null {
  const state = { tables };
  const assignment = pivotTableStyleAssignment(state, sheet, pivotIndex);
  const assigned = assignment ? customPivotTableStyleById(state, assignment.styleId) : null;
  const custom: CustomTableStyle | undefined =
    assigned ??
    (tables.customPivotTableStyles ?? []).find(
      (s) => s.label.toLocaleLowerCase() === name.toLocaleLowerCase(),
    );
  if (custom) return customStyleElementFormats(custom);
  return builtinPivotStyleElements(name, themeColors);
}

/** Build the style layer of `sheet`'s PivotTables. */
export function buildPivotStyleLayer(
  wb: PivotStyleReader,
  sheet: number,
  tables: Pick<TablesSlice, 'customPivotTableStyles' | 'pivotTableStyles'>,
): PivotStyleLayer {
  if (!wb.capabilities.pivotTables) return EMPTY_LAYER;
  const bounds = wb.pivotBounds(sheet);
  if (bounds.length === 0) return EMPTY_LAYER;
  const cellsByPivot = new Map<number, PivotLayoutCell[]>();
  for (const c of wb.pivotCells(sheet)) {
    const list = cellsByPivot.get(c.pivotIndex) ?? [];
    list.push({
      row: c.addr.row,
      col: c.addr.col,
      kind: c.kind,
      depth: c.depth,
      fieldName: c.fieldName,
      text:
        c.value.kind === 'text'
          ? c.value.value
          : c.value.kind === 'number'
            ? String(c.value.value)
            : '',
    });
    cellsByPivot.set(c.pivotIndex, list);
  }
  const themeColors = wb.getTheme()?.colors;
  const formats = new Map<string, CellFormat>();
  const buttons = new Set<string>();
  const outlines = new Map<string, PivotRowOutline>();
  const placeholders: PivotPlaceholder[] = [];
  const reports: PivotReportCells[] = [];
  for (const { pivotIndex, range } of bounds) {
    // The engine still projects a caption for a report with no fields.
    const cells = cellsByPivot.get(pivotIndex) ?? [];
    if (cells.every((c) => c.kind === PIVOT_CELL_KIND.Header)) {
      placeholders.push({
        pivotIndex,
        range: {
          sheet,
          r0: range.r0,
          c0: range.c0,
          r1: range.r0 + PIVOT_PLACEHOLDER_ROWS - 1,
          c1: range.c0 + PIVOT_PLACEHOLDER_COLS - 1,
        },
        name: wb.pivotTableName(sheet, pivotIndex) ?? '',
      });
      continue;
    }
    reports.push({ pivotIndex, range, cells });
    for (const label of pivotRowOutline(range, cellsByPivot.get(pivotIndex) ?? [])) {
      outlines.set(addrKey({ sheet, row: label.row, col: label.col }), label.outline);
    }
    const info = wb.getPivotStyleInfo(sheet, pivotIndex);
    if (!info) continue;
    const layout = pivotStyleRegions(range, cellsByPivot.get(pivotIndex) ?? [], info);
    for (const b of layout.buttons) buttons.add(addrKey({ sheet, ...b }));
    const elements = reportStyleElements(tables, sheet, pivotIndex, info.name, themeColors);
    if (!elements) continue;
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const format = composeStyleElements(
          elements,
          PIVOT_ELEMENT_ORDER,
          layout.regions,
          row,
          col,
        );
        if (format) formats.set(addrKey({ sheet, row, col }), format);
      }
    }
  }
  return {
    ranges: [
      ...reports.map((r) => ({ ...r.range, r0: pivotReportTop(r.range, r.cells) })),
      ...placeholders.map((p) => p.range),
    ],
    placeholders,
    reports,
    buttons,
    formatAt: (row, col) => formats.get(addrKey({ sheet, row, col })),
    outlineAt: (row, col) => outlines.get(addrKey({ sheet, row, col })),
  };
}

const layerCache = new WeakMap<
  object,
  { sheet: number; cells: unknown; tables: unknown; layer: PivotStyleLayer }
>();

/** The style layer of the sheet on show, rebuilt only when the sheet's cells
 *  or the style slice change (every report edit replaces the cell map). */
export function pivotStyleLayer(
  wb: PivotStyleReader | null,
  state: {
    data: { sheetIndex: number; cells: unknown };
    tables: Pick<TablesSlice, 'customPivotTableStyles' | 'pivotTableStyles'>;
  },
): PivotStyleLayer {
  if (!wb?.capabilities.pivotTables) return EMPTY_LAYER;
  const sheet = state.data.sheetIndex;
  const cached = layerCache.get(wb);
  if (
    cached &&
    cached.sheet === sheet &&
    cached.cells === state.data.cells &&
    cached.tables === state.tables
  ) {
    return cached.layer;
  }
  const layer = buildPivotStyleLayer(wb, sheet, state.tables);
  layerCache.set(wb, { sheet, cells: state.data.cells, tables: state.tables, layer });
  return layer;
}

/** The style layer last built for `sheet`, whatever has changed since; null
 *  when none was. A report edit reads the layout it replaced from here. */
export function lastPivotStyleLayer(wb: object, sheet: number): PivotStyleLayer | null {
  const cached = layerCache.get(wb);
  return cached && cached.sheet === sheet ? cached.layer : null;
}
