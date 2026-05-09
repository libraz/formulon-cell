import type { Range } from '../engine/types.js';
import type { SpreadsheetStore } from '../store/store.js';

/**
 * Sheet Views — the desktop spreadsheet's per-user filter / sort profiles. Each view
 * captures the current filter range, sort key, and frozen pane setup so
 * collaborators can switch perspectives without stepping on each other.
 *
 * The xlsx round-trip lives in the engine. This module is the UI-side
 * record + the pure helpers (capture / apply) that the store uses.
 * Persistence to disk is engine-gated; views captured today survive only
 * within the session unless the engine round-trips them.
 */

export interface SheetViewSort {
  /** A1 ref of the column being sorted (e.g. `"B"` or `"$C$1:$C$200"`). */
  column: string;
  direction: 'asc' | 'desc';
}

export interface SheetView {
  /** Unique id within the workbook — typically a UUID-ish string. */
  id: string;
  /** Display label, shown in the view-picker dropdown. */
  name: string;
  /** Sheet the view applies to (0-indexed). */
  sheet: number;
  /** Optional autofilter rect mirrored from `ui.filterRange`. */
  filterRange?: Range;
  /** Frozen-pane configuration. Default = neither row nor col frozen. */
  freeze?: { rows: number; cols: number };
  /** Active sort key (single-column, like the basic single-column sort UX). */
  sort?: SheetViewSort;
  /** Hidden rows / columns at capture time. Indices are 0-based. */
  hiddenRows?: readonly number[];
  hiddenCols?: readonly number[];
}

/** Snapshot input — the subset of store state a view should preserve. */
export interface SheetViewSnapshotInput {
  sheet: number;
  filterRange: Range | null;
  freezeRows: number;
  freezeCols: number;
  hiddenRows: ReadonlySet<number>;
  hiddenCols: ReadonlySet<number>;
  sort?: SheetViewSort;
}

/** Build a SheetView from the current store state. The id + name are
 *  caller-supplied so the picker can show user-friendly names. */
export function captureSheetView(
  id: string,
  name: string,
  input: SheetViewSnapshotInput,
): SheetView {
  const view: SheetView = {
    id,
    name,
    sheet: input.sheet,
  };
  if (input.filterRange) view.filterRange = input.filterRange;
  if (input.freezeRows > 0 || input.freezeCols > 0) {
    view.freeze = { rows: input.freezeRows, cols: input.freezeCols };
  }
  if (input.sort) view.sort = input.sort;
  if (input.hiddenRows.size > 0) {
    view.hiddenRows = [...input.hiddenRows].sort((a, b) => a - b);
  }
  if (input.hiddenCols.size > 0) {
    view.hiddenCols = [...input.hiddenCols].sort((a, b) => a - b);
  }
  return view;
}

/** Result of applying a view back to the store. The store mutator unpacks
 *  these fields and calls the existing layout / filter / freeze setters. */
export interface SheetViewPatch {
  sheet: number;
  filterRange: Range | null;
  freezeRows: number;
  freezeCols: number;
  hiddenRows: number[];
  hiddenCols: number[];
  sort: SheetViewSort | null;
}

/** Convert a captured view into a flat patch the store can replay. */
export function applySheetView(view: SheetView): SheetViewPatch {
  return {
    sheet: view.sheet,
    filterRange: view.filterRange ?? null,
    freezeRows: view.freeze?.rows ?? 0,
    freezeCols: view.freeze?.cols ?? 0,
    hiddenRows: [...(view.hiddenRows ?? [])],
    hiddenCols: [...(view.hiddenCols ?? [])],
    sort: view.sort ?? null,
  };
}

/** Find a view by id. Returns null when the id isn't tracked. */
export function findSheetView(views: readonly SheetView[], id: string): SheetView | null {
  return views.find((v) => v.id === id) ?? null;
}

/** Add a new view, replacing any existing entry with the same id. */
export function upsertSheetView(views: readonly SheetView[], next: SheetView): SheetView[] {
  const filtered = views.filter((v) => v.id !== next.id);
  filtered.push(next);
  return filtered;
}

/** Remove a view by id. Returns the same reference when nothing matched
 *  so consumers can short-circuit re-renders. */
export function removeSheetView(views: readonly SheetView[], id: string): readonly SheetView[] {
  const filtered = views.filter((v) => v.id !== id);
  if (filtered.length === views.length) return views;
  return filtered;
}

export type SheetViewStoreResult =
  | { ok: true; view: SheetView }
  | { ok: false; reason: 'not-found' | 'different-sheet' };

/** Capture the current store view settings and register them in the store. */
export function saveSheetView(
  store: SpreadsheetStore,
  id: string,
  name: string,
  sort?: SheetViewSort,
): SheetView {
  const state = store.getState();
  const view = captureSheetView(id, name, {
    sheet: state.data.sheetIndex,
    filterRange: state.ui.filterRange,
    freezeRows: state.layout.freezeRows,
    freezeCols: state.layout.freezeCols,
    hiddenRows: state.layout.hiddenRows,
    hiddenCols: state.layout.hiddenCols,
    sort,
  });
  store.setState((s) => {
    const next = s.sheetViews.views.filter((v) => v.id !== view.id);
    return { ...s, sheetViews: { ...s.sheetViews, views: [...next, view] } };
  });
  return view;
}

/** Apply a previously captured sheet view to the current sheet. */
export function activateSheetView(store: SpreadsheetStore, id: string): SheetViewStoreResult {
  const state = store.getState();
  const view = findSheetView(state.sheetViews.views, id);
  if (!view) return { ok: false, reason: 'not-found' };
  const patch = applySheetView(view);
  if (patch.sheet !== state.data.sheetIndex) return { ok: false, reason: 'different-sheet' };
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      freezeRows: Math.max(0, Math.floor(patch.freezeRows)),
      freezeCols: Math.max(0, Math.floor(patch.freezeCols)),
      hiddenRows: new Set(patch.hiddenRows),
      hiddenCols: new Set(patch.hiddenCols),
    },
    viewport: {
      ...s.viewport,
      rowStart: Math.max(Math.max(0, Math.floor(patch.freezeRows)), s.viewport.rowStart),
      colStart: Math.max(Math.max(0, Math.floor(patch.freezeCols)), s.viewport.colStart),
    },
    ui: { ...s.ui, filterRange: patch.filterRange },
    sheetViews: { ...s.sheetViews, activeViewId: id },
  }));
  return { ok: true, view };
}

/** Remove a stored sheet view and clear the active marker when needed. */
export function deleteSheetView(store: SpreadsheetStore, id: string): void {
  store.setState((s) => {
    const views = s.sheetViews.views.filter((v) => v.id !== id);
    if (views.length === s.sheetViews.views.length) return s;
    return {
      ...s,
      sheetViews: {
        views,
        activeViewId: s.sheetViews.activeViewId === id ? null : s.sheetViews.activeViewId,
      },
    };
  });
}
