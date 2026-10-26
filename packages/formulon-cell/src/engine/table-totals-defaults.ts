import type { TableColumnTotals } from '../commands/format-as-table.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Totals the desktop application gives a new total row: `label` in the first column and a sum on
 *  the last column that holds numbers in the data rows of `range` (which includes the total row). */
export function defaultTotals(
  wb: Pick<WorkbookHandle, 'getValue'>,
  range: Range,
  showHeader: boolean,
  columnCount: number,
  label: string,
): TableColumnTotals[] {
  const out: TableColumnTotals[] = Array.from({ length: columnCount }, () => ({
    function: 'none' as const,
    label: '',
    formula: '',
  }));
  const first = range.r0 + (showHeader ? 1 : 0);
  let sumColumn = -1;
  for (let col = columnCount - 1; col >= 0 && sumColumn < 0; col -= 1) {
    for (let row = first; row < range.r1; row += 1) {
      if (wb.getValue({ sheet: range.sheet, row, col: range.c0 + col }).kind === 'number') {
        sumColumn = col;
        break;
      }
    }
  }
  if (sumColumn !== 0 && out[0]) out[0] = { function: 'none', label, formula: '' };
  if (sumColumn >= 0 && out[sumColumn])
    out[sumColumn] = { function: 'sum', label: '', formula: '' };
  return out;
}
