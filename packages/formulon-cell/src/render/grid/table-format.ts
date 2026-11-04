import {
  isBandedCol,
  isBandedRow,
  isFirstCol,
  isHeaderRow,
  isLastCol,
  isTotalRow,
  type TableOverlay,
  tableForCell,
  tableStyleSwatch,
} from '../../commands/format-as-table.js';
import { pivotStyleLayer } from '../../engine/pivot-style.js';
import {
  composeStyleElements,
  type StyleElementFormats,
  type StyleElementRegion,
  TABLE_ELEMENT_ORDER,
} from '../../engine/table-style-elements.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { CellFormat } from '../../store/store.js';
import type { ViewState } from '../geometry.js';

/** The elements a gallery table style paints. */
function tableStyleElements(table: TableOverlay): StyleElementFormats {
  const swatch = tableStyleSwatch(table.style, table.color);
  const header = { fill: swatch.header, color: swatch.headerText, bold: true };
  return {
    headerRow: header,
    totalRow: header,
    firstRowStripe: { fill: swatch.band },
    firstColumnStripe: { fill: swatch.band },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
  };
}

export function tableCellFormat(
  table: TableOverlay,
  row: number,
  col: number,
): CellFormat | undefined {
  const { r0, r1, c0, c1 } = table.range;
  const regions: StyleElementRegion[] = [];
  if (isHeaderRow(table, row, col)) regions.push({ type: 'headerRow', r0, r1: r0, c0, c1 });
  if (isTotalRow(table, row, col)) regions.push({ type: 'totalRow', r0: r1, r1, c0, c1 });
  if (isBandedRow(table, row, col)) {
    regions.push({ type: 'firstRowStripe', r0: row, r1: row, c0, c1 });
  }
  if (isBandedCol(table, row, col)) {
    regions.push({ type: 'firstColumnStripe', r0, r1, c0: col, c1: col });
  }
  if (isFirstCol(table, row, col)) regions.push({ type: 'firstColumn', r0, r1, c0, c1: c0 });
  if (isLastCol(table, row, col)) regions.push({ type: 'lastColumn', r0, r1, c0: c1, c1 });
  if (regions.length === 0) return undefined;
  return composeStyleElements(tableStyleElements(table), TABLE_ELEMENT_ORDER, regions, row, col);
}

/** The format table and PivotTable styles paint at a cell of the sheet on
 *  show; the source the cell and border passes both read. */
export function styleElementFormatSource(
  state: Pick<ViewState, 'data' | 'tables'>,
  wb: WorkbookHandle | null,
): (row: number, col: number) => CellFormat | undefined {
  const pivots = pivotStyleLayer(wb, state);
  const sheet = state.data.sheetIndex;
  return (row, col) => {
    const table = tableForCell(state.tables.tables, sheet, row, col);
    const tableFmt = table ? tableCellFormat(table, row, col) : undefined;
    const pivotFmt = pivots.formatAt(row, col);
    return tableFmt && pivotFmt ? { ...tableFmt, ...pivotFmt } : (pivotFmt ?? tableFmt);
  };
}
