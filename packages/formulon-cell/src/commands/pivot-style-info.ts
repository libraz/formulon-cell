import type { PivotStyleDefinition } from '../engine/pivot-definition.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { PivotStyleInfo } from '../engine/workbook-handle-table-styles.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { type History, recordingHistory } from './history.js';

/** What a report with no style element is shown with. */
const DEFAULT_PIVOT_STYLE: PivotStyleDefinition = {
  name: 'PivotStyleLight16',
  showRowHeaders: true,
  showColHeaders: true,
  showRowStripes: false,
  showColStripes: false,
  showLastColumn: false,
};

/** The Design-tab toggles by command id, and the style fields they flip. */
export const PIVOT_STYLE_TOGGLES = {
  pivotRowHeaders: 'showRowHeaders',
  pivotColHeaders: 'showColHeaders',
  pivotRowStripes: 'showRowStripes',
  pivotColStripes: 'showColStripes',
} as const;

/** The toggle values of a report with no style element. */
export const DEFAULT_PIVOT_TOGGLES: Record<
  (typeof PIVOT_STYLE_TOGGLES)[keyof typeof PIVOT_STYLE_TOGGLES],
  boolean
> = {
  showRowHeaders: DEFAULT_PIVOT_STYLE.showRowHeaders,
  showColHeaders: DEFAULT_PIVOT_STYLE.showColHeaders,
  showRowStripes: DEFAULT_PIVOT_STYLE.showRowStripes,
  showColStripes: DEFAULT_PIVOT_STYLE.showColStripes,
};

export interface PivotStyleTarget {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History | null;
}

/**
 * Change a PivotTable's style name and/or its header and stripe toggles. Whatever the
 * patch leaves out keeps its value; the change is one undo step.
 */
export function updatePivotStyleInfo(
  target: PivotStyleTarget,
  pivot: { sheetIndex: number; pivotIndex: number },
  patch: PivotStyleInfo,
): boolean {
  const { store, workbook: wb, history } = target;
  const stored = wb.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex);
  const before = stored ?? DEFAULT_PIVOT_STYLE;
  const after: PivotStyleDefinition = { ...before, ...patch };
  if (!wb.setPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex, after)) return false;
  // The paint layer is keyed on the cell map, so a style change replaces it.
  const repaint = (): void => mutators.replaceCells(store, wb.cells(pivot.sheetIndex));
  repaint();
  const apply = (info: PivotStyleDefinition): void => {
    if (wb.setPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex, info)) repaint();
  };
  recordingHistory(history)?.push({ undo: () => apply(before), redo: () => apply(after) });
  return true;
}
