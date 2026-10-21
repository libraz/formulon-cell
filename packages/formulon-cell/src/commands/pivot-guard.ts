import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { selectionCoversRange } from '../store/selection-geometry.js';
import type { SelectionSlice } from '../store/types.js';

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

/** Tell the user a write would touch part of a PivotTable report. */
export function alertPivotEdit(wb: WorkbookHandle): void {
  guards.get(wb)?.sink(pivotGuardAlert(wb));
}

/** Reject a write over `range` that touches a PivotTable: alerts the user and
 *  returns true when the caller must leave the sheet untouched. */
export function blockPivotEdit(wb: WorkbookHandle, range: Range): boolean {
  if (!pivotOccupies(wb, range)) return false;
  alertPivotEdit(wb);
  return true;
}

/** How a selection meets the PivotTables of its sheet. */
export interface PivotSelectionPlan {
  /** Reports the selection contains in full; Delete removes these. */
  whole: { pivotIndex: number; range: Range }[];
  /** A report is only partly selected, so the edit must be refused. */
  partial: boolean;
}

export function planPivotSelection(
  wb: WorkbookHandle,
  selection: Pick<SelectionSlice, 'range' | 'extraRanges'>,
): PivotSelectionPlan {
  const ranges = [selection.range, ...(selection.extraRanges ?? [])];
  const plan: PivotSelectionPlan = { whole: [], partial: false };
  for (const bounds of wb.pivotBounds(selection.range.sheet)) {
    if (!ranges.some((range) => rangesIntersect(range, bounds.range))) {
      continue;
    }
    if (selectionCoversRange(selection, bounds.range)) plan.whole.push(bounds);
    else plan.partial = true;
  }
  return plan;
}

const rangesIntersect = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0;
