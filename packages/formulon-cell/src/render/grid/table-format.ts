import {
  isBandedRow,
  isFirstCol,
  isHeaderRow,
  isLastCol,
  isTotalRow,
  type TableOverlay,
  tableStyleSwatch,
} from '../../commands/format-as-table.js';
import type { CellFormat } from '../../store/store.js';

export function tableCellFormat(
  table: TableOverlay,
  row: number,
  col: number,
): CellFormat | undefined {
  const swatch = tableStyleSwatch(table.style, table.color);
  if (isHeaderRow(table, row, col)) {
    return { fill: swatch.header, color: swatch.headerText, bold: true };
  }
  if (isTotalRow(table, row, col)) {
    return { fill: swatch.header, color: swatch.headerText, bold: true };
  }
  const edge = isFirstCol(table, row, col) || isLastCol(table, row, col);
  if (isBandedRow(table, row, col)) {
    return edge ? { fill: swatch.band, bold: true } : { fill: swatch.band };
  }
  if (edge) return { bold: true };
  return undefined;
}
