import { addrKey } from '../../engine/address.js';
import type { CellValue, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';

type MaterializedCell = { value: CellValue; formula: string | null };

/** Materialize only the current logical copy range from the live workbook. */
export function materializedCopyCells(
  wb: WorkbookHandle,
  range: Range,
): Map<string, MaterializedCell> {
  const cells = new Map<string, MaterializedCell>();
  for (const cell of wb.cellsInRange(range)) {
    cells.set(addrKey(cell.addr), { value: cell.value, formula: cell.formula });
  }
  return cells;
}
