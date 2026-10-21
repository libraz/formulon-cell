import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** Stop-level alert shown when an edit would touch a PivotTable cell. */
export interface PivotGuardAlert {
  severity: 'stop';
  message: string;
}

type PivotGuardSink = (alert: PivotGuardAlert) => void;

const sinks = new WeakMap<WorkbookHandle, PivotGuardSink>();

/** Route `blockPivotEdit` alerts for `wb` to the host's alert surface. */
export function setPivotGuardSink(wb: WorkbookHandle, sink: PivotGuardSink | null): void {
  if (sink) sinks.set(wb, sink);
  else sinks.delete(wb);
}

export function pivotGuardAlert(wb: WorkbookHandle): PivotGuardAlert {
  return { severity: 'stop', message: wb.pivotCaptions().cannotChange };
}

/** Whether a PivotTable report occupies any cell of `range`. */
export function pivotOccupies(wb: WorkbookHandle, range: Range): boolean {
  return wb.pivotOverlapsRange(range);
}

/** Reject a write over `range` that touches a PivotTable: alerts the user and
 *  returns true when the caller must leave the sheet untouched. */
export function blockPivotEdit(wb: WorkbookHandle, range: Range): boolean {
  if (!pivotOccupies(wb, range)) return false;
  sinks.get(wb)?.(pivotGuardAlert(wb));
  return true;
}
