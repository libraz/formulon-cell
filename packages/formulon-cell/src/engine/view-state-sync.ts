import type { SpreadsheetStore } from '../store/store.js';
import type { SelectionSlice, SheetViewMemo } from '../store/types.js';
import type { Range } from './types.js';
import { SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type EngineSheetView = NonNullable<ReturnType<WorkbookHandle['getSheetView']>>;

/**
 * The visible sheet nearest `start` in desktop-spreadsheet order: the first
 * visible one at or to the right of it, else the nearest one on the left.
 * Returns -1 when every sheet is hidden.
 */
export function nearestVisibleSheet(
  count: number,
  hidden: ReadonlySet<number>,
  start: number,
): number {
  for (let i = Math.max(0, start); i < count; i += 1) if (!hidden.has(i)) return i;
  for (let i = Math.min(start, count) - 1; i >= 0; i -= 1) if (!hidden.has(i)) return i;
  return -1;
}

const a1Selection = (sheet: number): SelectionSlice => ({
  active: { sheet, row: 0, col: 0 },
  anchor: { sheet, row: 0, col: 0 },
  range: { sheet, r0: 0, c0: 0, r1: 0, c1: 0 },
});

const contains = (r: Range, row: number, col: number): boolean =>
  row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

const selectionFromEngine = (
  sheet: number,
  saved: ReturnType<WorkbookHandle['getSheetSelection']>,
): SelectionSlice => {
  if (!saved || saved.ranges.length === 0) return a1Selection(sheet);
  const active = { sheet, row: saved.activeRow, col: saved.activeCol };
  // The range holding the active cell becomes the primary; the rest are extras.
  const at = saved.ranges.findIndex((r) => contains(r, active.row, active.col));
  const primary = saved.ranges[Math.max(0, at)] as Range;
  const extraRanges = saved.ranges.filter((r) => r !== primary);
  return {
    active,
    anchor: { sheet, row: primary.r0, col: primary.c0 },
    range: primary,
    ...(extraRanges.length > 0 && { extraRanges }),
  };
};

/** Scrolling-pane start of one axis: the pane origin when frozen, else the sheet origin. */
const scrollStart = (frozen: number, topLeft: number, paneTopLeft: number): number =>
  frozen > 0 ? Math.max(paneTopLeft, topLeft + frozen) : topLeft;

const memoFromEngine = (
  sheet: number,
  view: EngineSheetView | null,
  saved: ReturnType<WorkbookHandle['getSheetSelection']>,
): SheetViewMemo => {
  const topRow = view?.topLeftRow ?? 0;
  const topCol = view?.topLeftCol ?? 0;
  const freezeRows = view?.freezeRows ?? 0;
  const freezeCols = view?.freezeCols ?? 0;
  return {
    selection: selectionFromEngine(sheet, saved),
    rowStart: scrollStart(freezeRows, topRow, view?.paneTopLeftRow ?? 0),
    colStart: scrollStart(freezeCols, topCol, view?.paneTopLeftCol ?? 0),
    freezeOriginRow: freezeRows > 0 ? topRow : 0,
    freezeOriginCol: freezeCols > 0 ? topCol : 0,
    workbookView: view?.viewMode ?? 'normal',
    zoomBeforePreview: null,
  };
};

/**
 * Open the store on the workbook's saved view: the active sheet plus each
 * sheet's selection, scroll origin and frozen-band origin, kept in the
 * per-sheet memos. Call after the workbook is bound and before the active
 * sheet's cells are hydrated. `followActiveSheet: false` leaves the current
 * sheet in place (a fixed viewport pins it).
 */
export function hydrateViewStateFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  opts: { followActiveSheet: boolean },
): void {
  const count = wb.sheetCount;
  const hidden = new Set<number>();
  const memos = new Map<number, SheetViewMemo>();
  for (let i = 0; i < count; i += 1) {
    const view = wb.getSheetView(i);
    if (view && view.visibility !== SheetVisibility.Visible) hidden.add(i);
    const saved = wb.getSheetSelection(i);
    if (view || saved) memos.set(i, memoFromEngine(i, view, saved));
  }
  const saved = opts.followActiveSheet ? wb.activeSheet() : null;
  const current = store.getState().data.sheetIndex;
  const target = saved === null ? current : nearestVisibleSheet(count, hidden, saved);
  const live = saved === null ? undefined : memos.get(target);
  memos.delete(saved === null ? current : target);
  store.setState((s) => ({
    ...s,
    data:
      target >= 0 && saved !== null ? { ...s.data, sheetIndex: target, cells: new Map() } : s.data,
    ui: { ...s.ui, sheetViewMemos: memos },
    ...(live && {
      selection: live.selection,
      layout: {
        ...s.layout,
        freezeOriginRow: live.freezeOriginRow,
        freezeOriginCol: live.freezeOriginCol,
      },
      viewport: { ...s.viewport, rowStart: live.rowStart, colStart: live.colStart },
    }),
  }));
}

const writeSheetView = (
  wb: WorkbookHandle,
  sheet: number,
  memo: Pick<
    SheetViewMemo,
    'selection' | 'rowStart' | 'colStart' | 'freezeOriginRow' | 'freezeOriginCol'
  >,
): void => {
  const { selection } = memo;
  wb.setSheetSelection(sheet, {
    activeRow: selection.active.row,
    activeCol: selection.active.col,
    ranges: [selection.range, ...(selection.extraRanges ?? [])],
  });
  const view = wb.getSheetView(sheet);
  const freezeRows = view?.freezeRows ?? 0;
  const freezeCols = view?.freezeCols ?? 0;
  const topRow = freezeRows > 0 ? memo.freezeOriginRow : memo.rowStart;
  const topCol = freezeCols > 0 ? memo.freezeOriginCol : memo.colStart;
  wb.setSheetTopLeft(sheet, topRow, topCol);
  if (freezeRows > 0 || freezeCols > 0) wb.setSheetPaneTopLeft(sheet, memo.rowStart, memo.colStart);
};

/**
 * Write the open-on sheet and every visited sheet's selection and scroll
 * origin to the engine so a save captures them. Sheets the user never visited
 * keep what the file carried.
 */
export function syncViewStateToEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  const s = store.getState();
  const active = s.data.sheetIndex;
  const count = wb.sheetCount;
  for (const [sheet, memo] of s.ui.sheetViewMemos ?? []) {
    if (sheet !== active && sheet < count) writeSheetView(wb, sheet, memo);
  }
  if (active < count) {
    writeSheetView(wb, active, {
      selection: s.selection,
      rowStart: s.viewport.rowStart,
      colStart: s.viewport.colStart,
      freezeOriginRow: s.layout.freezeOriginRow ?? 0,
      freezeOriginCol: s.layout.freezeOriginCol ?? 0,
    });
    if (wb.activeSheet() !== active) wb.setActiveSheet(active);
  }
}
