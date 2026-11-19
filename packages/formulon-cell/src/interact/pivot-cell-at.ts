import type { PivotCellInfo, PivotDefinition } from '../engine/pivot-definition.js';
import type { Addr, CellValue } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** The report layout entry at `addr`, or undefined when no report draws a cell there. */
export function pivotCellEntryAt(
  wb: WorkbookHandle,
  addr: Addr,
): ReturnType<WorkbookHandle['pivotCells']> extends Generator<infer E> ? E | undefined : never {
  for (const entry of wb.pivotCells(addr.sheet)) {
    if (entry.addr.row === addr.row && entry.addr.col === addr.col) return entry;
  }
  return undefined;
}

/** The cell's value as its report draws it; the sheet value may not carry the item text yet. */
export function pivotCellValueAt(wb: WorkbookHandle, addr: Addr): CellValue {
  return pivotCellEntryAt(wb, addr)?.value ?? wb.getValue(addr);
}

/** The report field a `pivotCellInfo` names, by its cache field. */
export function pivotFieldOfCell(
  definition: PivotDefinition | null | undefined,
  info: Pick<PivotCellInfo, 'fieldIndex'>,
): PivotDefinition['fields'][number] | undefined {
  return definition?.fields.find((f) => f.cacheFieldIndex === info.fieldIndex);
}
