import {
  approximateTableStyleElements,
  type CustomTableStyle,
  isBandedCol,
  isBandedRow,
  isFirstCol,
  isHeaderRow,
  isLastCol,
  isTotalRow,
  overlayStyleName,
  type TableOverlay,
  tableForCell,
} from '../../commands/format-as-table.js';
import { builtinTableStyleElements } from '../../engine/builtin-table-styles.js';
import { pivotStyleLayer } from '../../engine/pivot-style.js';
import {
  composeStyleElements,
  type StyleElementRegion,
  TABLE_ELEMENT_ORDER,
} from '../../engine/table-style-elements.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { CellFormat } from '../../store/store.js';
import type { ViewState } from '../geometry.js';

/** A body cell: inside the table, below the header and above the total row. */
const inBody = (table: TableOverlay, row: number, col: number): boolean =>
  row >= table.range.r0 &&
  row <= table.range.r1 &&
  col >= table.range.c0 &&
  col <= table.range.c1 &&
  !isHeaderRow(table, row, col) &&
  !isTotalRow(table, row, col);

/** The format a table paints at (row, col): a built-in style from its own
 *  definition under the workbook theme, a custom one from its elements, else
 *  an approximation by family and hue. */
export function tableCellFormat(
  table: TableOverlay,
  row: number,
  col: number,
  themeColors?: readonly number[],
  customs: readonly CustomTableStyle[] = [],
): CellFormat | undefined {
  const { r0, r1, c0, c1 } = table.range;
  const name = overlayStyleName(table);
  const builtin = builtinTableStyleElements(name, themeColors);
  const custom = builtin
    ? undefined
    : customs.find((c) => c.label.toLocaleLowerCase() === name.toLocaleLowerCase());
  const sizes = custom?.stripeSizes;
  const regions: StyleElementRegion[] = [{ type: 'wholeTable', r0, r1, c0, c1 }];
  const header = isHeaderRow(table, row, col);
  const total = isTotalRow(table, row, col);
  if (header) regions.push({ type: 'headerRow', r0, r1: r0, c0, c1 });
  if (total) regions.push({ type: 'totalRow', r0: r1, r1, c0, c1 });
  if (header) {
    if (col === c0) regions.push({ type: 'firstHeaderCell', r0, r1: r0, c0, c1: c0 });
    if (col === c1) regions.push({ type: 'lastHeaderCell', r0, r1: r0, c0: c1, c1 });
  }
  if (total) {
    if (col === c0) regions.push({ type: 'firstTotalCell', r0: r1, r1, c0, c1: c0 });
    if (col === c1) regions.push({ type: 'lastTotalCell', r0: r1, r1, c0: c1, c1 });
  }
  if (table.banded && inBody(table, row, col)) {
    const type = isBandedRow(table, row, col, {
      first: sizes?.firstRowStripe,
      second: sizes?.secondRowStripe,
    })
      ? 'firstRowStripe'
      : 'secondRowStripe';
    regions.push({ type, r0: row, r1: row, c0, c1 });
  }
  if (table.colStripes && inBody(table, row, col)) {
    const type = isBandedCol(table, row, col, {
      first: sizes?.firstColumnStripe,
      second: sizes?.secondColumnStripe,
    })
      ? 'firstColumnStripe'
      : 'secondColumnStripe';
    regions.push({ type, r0, r1, c0: col, c1: col });
  }
  if (isFirstCol(table, row, col)) regions.push({ type: 'firstColumn', r0, r1, c0, c1: c0 });
  if (isLastCol(table, row, col)) regions.push({ type: 'lastColumn', r0, r1, c0: c1, c1 });
  if (name === '') return undefined;
  const elements =
    builtin ?? custom?.elements ?? approximateTableStyleElements(table.style, table.color);
  return composeStyleElements(elements, TABLE_ELEMENT_ORDER, regions, row, col);
}

/** The format table and PivotTable styles paint at a cell of the sheet on
 *  show; the source the cell and border passes both read. */
export function styleElementFormatSource(
  state: Pick<ViewState, 'data' | 'tables'>,
  wb: WorkbookHandle | null,
): (row: number, col: number) => CellFormat | undefined {
  const pivots = pivotStyleLayer(wb, state);
  const sheet = state.data.sheetIndex;
  const themeColors = wb?.getTheme?.()?.colors;
  return (row, col) => {
    const table = tableForCell(state.tables.tables, sheet, row, col);
    const tableFmt = table
      ? tableCellFormat(table, row, col, themeColors, state.tables.customTableStyles)
      : undefined;
    const pivotFmt = pivots.formatAt(row, col);
    return tableFmt && pivotFmt ? { ...tableFmt, ...pivotFmt } : (pivotFmt ?? tableFmt);
  };
}
