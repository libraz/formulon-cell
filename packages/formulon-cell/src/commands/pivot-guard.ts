import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangesIntersect, selectionCoversRange } from '../store/selection-geometry.js';
import type { SelectionSlice } from '../store/types.js';

/** Stop-level alert shown when an edit would touch a PivotTable cell. */
export interface PivotGuardAlert {
  severity: 'stop';
  message: string;
}

type PivotGuardSink = (alert: PivotGuardAlert) => void;

/** Which gesture hit the report: a paste gets its own, longer wording. */
export type PivotEditKind = 'edit' | 'paste';

export interface PivotGuardMessages {
  edit: string;
  paste: string;
}

const ENGLISH_MESSAGES: PivotGuardMessages = {
  edit: 'Cannot change this part of a PivotTable report.',
  paste:
    "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again.",
};

const guards = new WeakMap<
  WorkbookHandle,
  { sink: PivotGuardSink; messages: PivotGuardMessages }
>();

/** Route `blockPivotEdit` alerts for `wb` to the host's alert surface, worded
 *  in the UI language. */
export function setPivotGuardSink(
  wb: WorkbookHandle,
  sink: PivotGuardSink | null,
  messages: PivotGuardMessages = ENGLISH_MESSAGES,
): void {
  if (sink) guards.set(wb, { sink, messages });
  else guards.delete(wb);
}

export function pivotGuardAlert(wb: WorkbookHandle, kind: PivotEditKind = 'edit'): PivotGuardAlert {
  return { severity: 'stop', message: (guards.get(wb)?.messages ?? ENGLISH_MESSAGES)[kind] };
}

/** Whether a PivotTable report occupies any cell of `range`. */
export function pivotOccupies(wb: WorkbookHandle, range: Range): boolean {
  return wb.pivotOverlapsRange(range);
}

/** Tell the user a write would touch part of a PivotTable report. */
export function alertPivotEdit(wb: WorkbookHandle, kind: PivotEditKind = 'edit'): void {
  guards.get(wb)?.sink(pivotGuardAlert(wb, kind));
}

/** Reject a write over `range` that touches a PivotTable: alerts the user and
 *  returns true when the caller must leave the sheet untouched. */
export function blockPivotEdit(
  wb: WorkbookHandle,
  range: Range,
  kind: PivotEditKind = 'edit',
): boolean {
  if (!pivotOccupies(wb, range)) return false;
  alertPivotEdit(wb, kind);
  return true;
}

/** Whether inserting `count` rows or columns before line `at`, or deleting
 *  `count` of them from `at`, cuts through a report on `sheet`. An insertion
 *  at a report's first line only moves it. */
export function axisEditCrossesPivot(
  wb: WorkbookHandle,
  sheet: number,
  axis: 'row' | 'col',
  kind: 'insert' | 'delete',
  at: number,
  count: number,
): boolean {
  if (!wb.capabilities?.pivotTables) return false;
  return wb.pivotBounds(sheet).some(({ range }) => {
    const first = axis === 'row' ? range.r0 : range.c0;
    const last = axis === 'row' ? range.r1 : range.c1;
    return kind === 'insert' ? at > first && at <= last : at <= last && at + count - 1 >= first;
  });
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
