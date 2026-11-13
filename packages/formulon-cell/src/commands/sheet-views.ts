import type { Range } from '../engine/types.js';
import { frozenColEnd, frozenRowEnd } from '../store/freeze-band.js';
import { cloneRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, ValueFilterCriteria } from '../store/store.js';
import type { History } from './history.js';

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
  /** AutoFilter ranges mirrored from `ui.filterRanges`. */
  filterRanges?: Range[];
  /** Value-filter criteria captured with the view so Reapply and dropdown
   *  state survive switching between sheet views. */
  filterCriteria?: readonly ValueFilterCriteria[];
  /** Frozen-pane configuration. Default = neither row nor col frozen. */
  freeze?: { rows: number; cols: number; originRow?: number; originCol?: number };
  /** Active sort key (single-column, like the basic single-column sort UX). */
  sort?: SheetViewSort;
  /** Hidden rows / columns at capture time. Indices are 0-based. */
  hiddenRows?: readonly number[];
  hiddenCols?: readonly number[];
}

/** Snapshot input — the subset of store state a view should preserve. */
export interface SheetViewSnapshotInput {
  sheet: number;
  filterRanges: readonly Range[];
  filterCriteria?: readonly ValueFilterCriteria[];
  freezeRows: number;
  freezeCols: number;
  freezeOriginRow?: number;
  freezeOriginCol?: number;
  hiddenRows: ReadonlySet<number>;
  hiddenCols: ReadonlySet<number>;
  sort?: SheetViewSort;
}

const cloneFilterCriteria = (criteria: readonly ValueFilterCriteria[]): ValueFilterCriteria[] =>
  criteria.map((criterion) => ({
    range: { ...criterion.range },
    byCol: criterion.byCol,
    hiddenValues: [...criterion.hiddenValues],
  }));

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
  if (input.filterRanges.length) view.filterRanges = input.filterRanges.map(cloneRange);
  if (input.filterCriteria?.length) view.filterCriteria = cloneFilterCriteria(input.filterCriteria);
  if (input.freezeRows > 0 || input.freezeCols > 0) {
    view.freeze = {
      rows: input.freezeRows,
      cols: input.freezeCols,
      originRow: input.freezeOriginRow ?? 0,
      originCol: input.freezeOriginCol ?? 0,
    };
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
  filterRanges: Range[];
  filterCriteria: ValueFilterCriteria[];
  freezeRows: number;
  freezeCols: number;
  freezeOriginRow: number;
  freezeOriginCol: number;
  hiddenRows: number[];
  hiddenCols: number[];
  sort: SheetViewSort | null;
}

/** Convert a captured view into a flat patch the store can replay. */
export function applySheetView(view: SheetView): SheetViewPatch {
  return {
    sheet: view.sheet,
    filterRanges: (view.filterRanges ?? []).map(cloneRange),
    filterCriteria: cloneFilterCriteria(view.filterCriteria ?? []),
    freezeRows: view.freeze?.rows ?? 0,
    freezeCols: view.freeze?.cols ?? 0,
    freezeOriginRow: view.freeze?.originRow ?? 0,
    freezeOriginCol: view.freeze?.originCol ?? 0,
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

/** Select value of Default, the sheet's own view outside every sheet view. */
export const DEFAULT_SHEET_VIEW_ID = 'current';
/** `activeViewId` of the view New opened and Keep has not named yet. */
export const TEMPORARY_SHEET_VIEW_ID = 'temporary';

interface SheetViewsSnapshot {
  views: SheetView[];
  activeViewId: string | null;
  temporary: SheetView | null;
  defaultView: SheetView | null;
  /** The view settings on screen, so undo puts back what a switch replaced. */
  shown: SheetView;
}

const cloneSheetView = (view: SheetView): SheetView => ({
  ...view,
  ...(view.filterRanges ? { filterRanges: view.filterRanges.map(cloneRange) } : {}),
  ...(view.filterCriteria ? { filterCriteria: cloneFilterCriteria(view.filterCriteria) } : {}),
  ...(view.freeze ? { freeze: { ...view.freeze } } : {}),
  ...(view.sort ? { sort: { ...view.sort } } : {}),
  ...(view.hiddenRows ? { hiddenRows: [...view.hiddenRows] } : {}),
  ...(view.hiddenCols ? { hiddenCols: [...view.hiddenCols] } : {}),
});

const cloneOrNull = (view: SheetView | null): SheetView | null =>
  view ? cloneSheetView(view) : null;

/** The current sheet's view settings as a view with the given id and name. */
const captureShown = (
  store: SpreadsheetStore,
  id: string,
  name: string,
  sort?: SheetViewSort,
): SheetView => {
  const state = store.getState();
  return captureSheetView(id, name, {
    sheet: state.data.sheetIndex,
    filterRanges: state.ui.filterRanges,
    filterCriteria: state.ui.filterCriteria,
    freezeRows: state.layout.freezeRows,
    freezeCols: state.layout.freezeCols,
    freezeOriginRow: state.layout.freezeOriginRow,
    freezeOriginCol: state.layout.freezeOriginCol,
    hiddenRows: state.layout.hiddenRows,
    hiddenCols: state.layout.hiddenCols,
    sort,
  });
};

/** Puts a view's settings on screen. */
const showSettings = (store: SpreadsheetStore, view: SheetView): void => {
  const patch = applySheetView(view);
  store.setState((s) => {
    const layout = {
      ...s.layout,
      freezeRows: Math.max(0, Math.floor(patch.freezeRows)),
      freezeCols: Math.max(0, Math.floor(patch.freezeCols)),
      freezeOriginRow: Math.max(0, Math.floor(patch.freezeOriginRow)),
      freezeOriginCol: Math.max(0, Math.floor(patch.freezeOriginCol)),
      hiddenRows: new Set(patch.hiddenRows),
      hiddenCols: new Set(patch.hiddenCols),
    };
    return {
      ...s,
      layout,
      viewport: {
        ...s.viewport,
        rowStart: Math.max(frozenRowEnd(layout), s.viewport.rowStart),
        colStart: Math.max(frozenColEnd(layout), s.viewport.colStart),
      },
      ui: { ...s.ui, filterRanges: patch.filterRanges, filterCriteria: patch.filterCriteria },
    };
  });
};

const captureSheetViewsSnapshot = (store: SpreadsheetStore): SheetViewsSnapshot => {
  const { sheetViews } = store.getState();
  return {
    views: sheetViews.views.map(cloneSheetView),
    activeViewId: sheetViews.activeViewId,
    temporary: cloneOrNull(sheetViews.temporary),
    defaultView: cloneOrNull(sheetViews.defaultView),
    shown: captureShown(store, '', ''),
  };
};

const applySheetViewsSnapshot = (store: SpreadsheetStore, snapshot: SheetViewsSnapshot): void => {
  showSettings(store, snapshot.shown);
  store.setState((s) => ({
    ...s,
    sheetViews: {
      views: snapshot.views.map(cloneSheetView),
      activeViewId: snapshot.activeViewId,
      temporary: cloneOrNull(snapshot.temporary),
      defaultView: cloneOrNull(snapshot.defaultView),
    },
  }));
};

const sameSheetViewsSnapshot = (a: SheetViewsSnapshot, b: SheetViewsSnapshot): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

export function recordSheetViewsChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => T,
): T {
  if (!history || history.isReplaying()) return mutate();
  const before = captureSheetViewsSnapshot(store);
  const result = mutate();
  const after = captureSheetViewsSnapshot(store);
  if (!sameSheetViewsSnapshot(before, after)) {
    history.push({
      undo: () => applySheetViewsSnapshot(store, before),
      redo: () => applySheetViewsSnapshot(store, after),
    });
  }
  return result;
}

/** Keeps the shown view's settings as it is left: Default's are held, a named view's are
 *  updated, and the temporary view is discarded. */
const leaveShownView = (store: SpreadsheetStore): void => {
  const { activeViewId, views } = store.getState().sheetViews;
  if (activeViewId === TEMPORARY_SHEET_VIEW_ID) {
    store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, temporary: null } }));
    return;
  }
  if (activeViewId === null) {
    const held = captureShown(store, DEFAULT_SHEET_VIEW_ID, '');
    store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, defaultView: held } }));
    return;
  }
  const named = findSheetView(views, activeViewId);
  if (!named) return;
  const updated = captureShown(store, named.id, named.name, named.sort);
  store.setState((s) => ({
    ...s,
    sheetViews: {
      ...s.sheetViews,
      views: s.sheetViews.views.map((v) => (v.id === named.id ? updated : v)),
    },
  }));
};

const markShown = (store: SpreadsheetStore, activeViewId: string | null): void => {
  store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, activeViewId } }));
};

/** The first free id and numbered name for a new view of the current sheet. */
export function nextSheetViewName(
  store: SpreadsheetStore,
  prefix: string,
): { id: string; name: string } {
  const { sheetViews, data } = store.getState();
  const taken = new Set(sheetViews.views.map((v) => v.id));
  let n = 1;
  while (taken.has(`view-${data.sheetIndex}-${n}`)) n += 1;
  return { id: `view-${data.sheetIndex}-${n}`, name: `${prefix} ${n}` };
}

/** Capture the current store view settings and register them in the store. */
export function saveSheetView(
  store: SpreadsheetStore,
  id: string,
  name: string,
  sort?: SheetViewSort,
  history: History | null = null,
): SheetView {
  const view = captureShown(store, id, name, sort);
  recordSheetViewsChange(history, store, () => {
    store.setState((s) => {
      const next = s.sheetViews.views.filter((v) => v.id !== view.id);
      return { ...s, sheetViews: { ...s.sheetViews, views: [...next, view] } };
    });
  });
  return view;
}

/** Switch to a named view of the current sheet, keeping the settings of the view left. */
export function activateSheetView(
  store: SpreadsheetStore,
  id: string,
  history: History | null = null,
): SheetViewStoreResult {
  const state = store.getState();
  const view = findSheetView(state.sheetViews.views, id);
  if (!view) return { ok: false, reason: 'not-found' };
  if (view.sheet !== state.data.sheetIndex) return { ok: false, reason: 'different-sheet' };
  if (state.sheetViews.activeViewId === id) return { ok: true, view };
  recordSheetViewsChange(history, store, () => {
    leaveShownView(store);
    showSettings(store, view);
    markShown(store, id);
  });
  return { ok: true, view };
}

/** Open a temporary view holding the current settings in place of the shown view. */
export function newSheetView(store: SpreadsheetStore, history: History | null = null): SheetView {
  return recordSheetViewsChange(history, store, () => {
    leaveShownView(store);
    const temporary = captureShown(store, TEMPORARY_SHEET_VIEW_ID, '');
    store.setState((s) => ({
      ...s,
      sheetViews: { ...s.sheetViews, temporary, activeViewId: TEMPORARY_SHEET_VIEW_ID },
    }));
    return temporary;
  });
}

/** Keep the open temporary view as a named view; null when no temporary view is open. */
export function keepSheetView(
  store: SpreadsheetStore,
  id: string,
  name: string,
  history: History | null = null,
): SheetView | null {
  if (store.getState().sheetViews.activeViewId !== TEMPORARY_SHEET_VIEW_ID) return null;
  const view = captureShown(store, id, name);
  recordSheetViewsChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      sheetViews: {
        ...s.sheetViews,
        views: [...s.sheetViews.views.filter((v) => v.id !== id), view],
        temporary: null,
        activeViewId: id,
      },
    }));
  });
  return view;
}

/** Leave the shown sheet view for Default, putting Default's held settings back. */
export function exitSheetView(store: SpreadsheetStore, history: History | null = null): void {
  if (store.getState().sheetViews.activeViewId === null) return;
  recordSheetViewsChange(history, store, () => {
    leaveShownView(store);
    const { defaultView } = store.getState().sheetViews;
    if (defaultView && defaultView.sheet === store.getState().data.sheetIndex) {
      showSettings(store, defaultView);
    }
    store.setState((s) => ({
      ...s,
      sheetViews: { ...s.sheetViews, defaultView: null, activeViewId: null },
    }));
  });
}

/** Why a name cannot be given to a view of the current sheet, or null when it can. */
export type SheetViewNameProblem = 'invalid' | 'taken';

export function sheetViewNameProblem(
  store: SpreadsheetStore,
  name: string,
  exceptId?: string,
): SheetViewNameProblem | null {
  if (name.trim() === '' || name.includes(';')) return 'invalid';
  const { sheetViews, data } = store.getState();
  const folded = name.trim().toLocaleLowerCase();
  const taken = sheetViews.views.some(
    (v) =>
      v.sheet === data.sheetIndex && v.id !== exceptId && v.name.toLocaleLowerCase() === folded,
  );
  return taken ? 'taken' : null;
}

/** Rename a saved view; refused names leave it as it is. */
export function renameSheetView(
  store: SpreadsheetStore,
  id: string,
  name: string,
  history: History | null = null,
): SheetViewNameProblem | 'not-found' | null {
  if (!findSheetView(store.getState().sheetViews.views, id)) return 'not-found';
  const problem = sheetViewNameProblem(store, name, id);
  if (problem) return problem;
  recordSheetViewsChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      sheetViews: {
        ...s.sheetViews,
        views: s.sheetViews.views.map((v) => (v.id === id ? { ...v, name: name.trim() } : v)),
      },
    }));
  });
  return null;
}

/** The settings a view holds now: the screen's for the shown view, else its stored ones.
 *  `DEFAULT_SHEET_VIEW_ID` reads Default. */
const settingsOf = (store: SpreadsheetStore, id: string): SheetView | null => {
  const { sheetViews } = store.getState();
  const shown = (sheetViews.activeViewId ?? DEFAULT_SHEET_VIEW_ID) === id;
  if (shown) return captureShown(store, id, '');
  if (id === DEFAULT_SHEET_VIEW_ID) return sheetViews.defaultView ?? captureShown(store, id, '');
  return findSheetView(sheetViews.views, id);
};

/** Copy a saved view or Default into a new saved view named `name`; null when the source is
 *  missing or the name is refused. */
export function duplicateSheetView(
  store: SpreadsheetStore,
  sourceId: string,
  id: string,
  name: string,
  history: History | null = null,
): SheetView | null {
  const source = settingsOf(store, sourceId);
  if (!source || sheetViewNameProblem(store, name)) return null;
  const copy: SheetView = { ...cloneSheetView(source), id, name };
  recordSheetViewsChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      sheetViews: { ...s.sheetViews, views: [...s.sheetViews.views, copy] },
    }));
  });
  return copy;
}

/** Remove a stored sheet view; removing the shown one returns to Default. */
export function deleteSheetView(
  store: SpreadsheetStore,
  id: string,
  history: History | null = null,
): void {
  const { views, activeViewId, defaultView } = store.getState().sheetViews;
  if (!findSheetView(views, id)) return;
  recordSheetViewsChange(history, store, () => {
    const shown = activeViewId === id;
    if (shown && defaultView && defaultView.sheet === store.getState().data.sheetIndex) {
      showSettings(store, defaultView);
    }
    store.setState((s) => ({
      ...s,
      sheetViews: {
        ...s.sheetViews,
        views: s.sheetViews.views.filter((v) => v.id !== id),
        ...(shown ? { activeViewId: null, defaultView: null } : {}),
      },
    }));
  });
}
