import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** Stop-level alert shown when an edit would touch a PivotTable cell. */
export interface PivotGuardAlert {
  severity: 'stop';
  message: string;
}

type PivotGuardSink = (alert: PivotGuardAlert) => void;

const ENGLISH_MESSAGE = 'Cannot change this part of a PivotTable report.';

const guards = new WeakMap<WorkbookHandle, { sink: PivotGuardSink; message: string }>();

/** Route `blockPivotEdit` alerts for `wb` to the host's alert surface, worded
 *  in the UI language. */
export function setPivotGuardSink(
  wb: WorkbookHandle,
  sink: PivotGuardSink | null,
  message = ENGLISH_MESSAGE,
): void {
  if (sink) guards.set(wb, { sink, message });
  else guards.delete(wb);
}

export function pivotGuardAlert(wb: WorkbookHandle): PivotGuardAlert {
  return { severity: 'stop', message: guards.get(wb)?.message ?? ENGLISH_MESSAGE };
}

/** Whether a PivotTable report occupies any cell of `range`. */
export function pivotOccupies(wb: WorkbookHandle, range: Range): boolean {
  return wb.pivotOverlapsRange(range);
}

/** Reject a write over `range` that touches a PivotTable: alerts the user and
 *  returns true when the caller must leave the sheet untouched. */
export function blockPivotEdit(wb: WorkbookHandle, range: Range): boolean {
  if (!pivotOccupies(wb, range)) return false;
  guards.get(wb)?.sink(pivotGuardAlert(wb));
  return true;
}
