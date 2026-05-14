import {
  isBandedRow,
  isHeaderRow,
  isTotalRow,
  type TableOverlay,
} from '../../commands/format-as-table.js';
import type { CellFormat } from '../../store/store.js';

export function tableCellFormat(
  table: TableOverlay,
  row: number,
  col: number,
): CellFormat | undefined {
  if (isHeaderRow(table, row, col)) {
    return {
      fill: table.style === 'dark' ? '#1f4e78' : table.style === 'light' ? '#d9eaf7' : '#5b9bd5',
      color: table.style === 'light' ? '#1f1f1f' : '#ffffff',
      bold: true,
    };
  }
  if (isTotalRow(table, row, col)) {
    return {
      fill: table.style === 'dark' ? '#385723' : table.style === 'light' ? '#e2f0d9' : '#a9d18e',
      color: '#1f1f1f',
      bold: true,
    };
  }
  if (isBandedRow(table, row, col)) {
    return {
      fill: table.style === 'dark' ? '#d9e2f3' : table.style === 'light' ? '#f3f8fc' : '#ddebf7',
    };
  }
  return undefined;
}
