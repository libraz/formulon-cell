import type { Addr, Range } from '../engine/types.js';
import type { SpreadsheetStore, State } from '../store/store.js';

interface SelectionSnapshot {
  readonly active: Addr;
  readonly anchor: Addr;
  readonly range: Range;
  readonly extraRanges: readonly Range[];
}

interface TabNavigationSession {
  laneRow: number | null;
  returnOrigin: Addr | null;
  expectedSelection: SelectionSnapshot;
  sheetIndex: number;
  hiddenRows: State['layout']['hiddenRows'];
  hiddenCols: State['layout']['hiddenCols'];
  shownRows: State['layout']['shownRows'];
  defaultRowsHidden: State['layout']['defaultRowsHidden'];
  merges: State['merges'];
  guardDepth: number;
  unsubscribe: (() => void) | null;
}

export interface TabNavigationMemory {
  readonly laneRow: number | null;
  readonly returnOrigin: Addr | null;
}

const sessions = new WeakMap<SpreadsheetStore, TabNavigationSession>();

const copyAddr = (addr: Addr): Addr => ({ sheet: addr.sheet, row: addr.row, col: addr.col });
const copyRange = (range: Range): Range => ({
  sheet: range.sheet,
  r0: range.r0,
  c0: range.c0,
  r1: range.r1,
  c1: range.c1,
});

const snapshotSelection = (selection: State['selection']): SelectionSnapshot => ({
  active: copyAddr(selection.active),
  anchor: copyAddr(selection.anchor),
  range: copyRange(selection.range),
  extraRanges: (selection.extraRanges ?? []).map(copyRange),
});

const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

const sameRange = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;

const sameSelection = (a: SelectionSnapshot, b: State['selection']): boolean =>
  sameAddr(a.active, b.active) &&
  sameAddr(a.anchor, b.anchor) &&
  sameRange(a.range, b.range) &&
  a.extraRanges.length === (b.extraRanges?.length ?? 0) &&
  a.extraRanges.every((range, index) => {
    const other = b.extraRanges?.[index];
    return other !== undefined && sameRange(range, other);
  });

function resetMemory(session: TabNavigationSession): void {
  session.laneRow = null;
  session.returnOrigin = null;
}

function createSession(store: SpreadsheetStore): TabNavigationSession {
  const state = store.getState();
  const session: TabNavigationSession = {
    laneRow: null,
    returnOrigin: null,
    expectedSelection: snapshotSelection(state.selection),
    sheetIndex: state.data.sheetIndex,
    hiddenRows: state.layout.hiddenRows,
    hiddenCols: state.layout.hiddenCols,
    shownRows: state.layout.shownRows,
    defaultRowsHidden: state.layout.defaultRowsHidden,
    merges: state.merges,
    guardDepth: 0,
    unsubscribe: null,
  };
  session.unsubscribe = store.subscribe((next) => {
    const selectionChanged = !sameSelection(session.expectedSelection, next.selection);
    if (selectionChanged && session.guardDepth === 0) resetMemory(session);

    if (
      next.data.sheetIndex !== session.sheetIndex ||
      next.layout.hiddenRows !== session.hiddenRows ||
      next.layout.hiddenCols !== session.hiddenCols ||
      next.layout.shownRows !== session.shownRows ||
      next.layout.defaultRowsHidden !== session.defaultRowsHidden ||
      next.merges !== session.merges
    ) {
      resetMemory(session);
    }

    session.expectedSelection = snapshotSelection(next.selection);
    session.sheetIndex = next.data.sheetIndex;
    session.hiddenRows = next.layout.hiddenRows;
    session.hiddenCols = next.layout.hiddenCols;
    session.shownRows = next.layout.shownRows;
    session.defaultRowsHidden = next.layout.defaultRowsHidden;
    session.merges = next.merges;
  });
  return session;
}

function ensureSession(store: SpreadsheetStore): TabNavigationSession {
  const existing = sessions.get(store);
  if (existing) return existing;
  const session = createSession(store);
  sessions.set(store, session);
  return session;
}

/** Attach the ephemeral navigation listener for a mounted store. */
export function attachTabNavigationSession(store: SpreadsheetStore): () => void {
  disposeTabNavigationSession(store);
  const session = ensureSession(store);
  return () => {
    if (sessions.get(store) === session) disposeTabNavigationSession(store);
  };
}

/** Remove the listener and discard all per-store navigation memory. */
export function disposeTabNavigationSession(store: SpreadsheetStore): void {
  const session = sessions.get(store);
  if (!session) return;
  sessions.delete(store);
  resetMemory(session);
  session.unsubscribe?.();
  session.unsubscribe = null;
}

/** Clear both parts of the transient Mac Tab-entry state, if it is attached. */
export function resetTabNavigationSession(store: SpreadsheetStore): void {
  const session = sessions.get(store);
  if (session) resetMemory(session);
}

/** Clear only Return's origin while preserving the horizontal Tab lane. */
export function clearTabReturnOrigin(store: SpreadsheetStore): void {
  const session = sessions.get(store);
  if (session) session.returnOrigin = null;
}

export function tabNavigationMemoryFor(store: SpreadsheetStore): TabNavigationMemory | null {
  const session = sessions.get(store);
  if (!session) return null;
  return {
    laneRow: session.laneRow,
    returnOrigin: session.returnOrigin ? copyAddr(session.returnOrigin) : null,
  };
}

export function setTabNavigationMemory(
  store: SpreadsheetStore,
  patch: { readonly laneRow?: number | null; readonly returnOrigin?: Addr | null },
): void {
  const session = ensureSession(store);
  if (patch.laneRow !== undefined) session.laneRow = patch.laneRow;
  if (patch.returnOrigin !== undefined) {
    session.returnOrigin = patch.returnOrigin ? copyAddr(patch.returnOrigin) : null;
  }
}

/** Keep the selection subscriber from treating a session-owned move as external. */
export function withTabNavigationSelectionMutation<T>(store: SpreadsheetStore, mutate: () => T): T {
  const session = ensureSession(store);
  session.guardDepth += 1;
  try {
    return mutate();
  } finally {
    session.expectedSelection = snapshotSelection(store.getState().selection);
    session.guardDepth -= 1;
  }
}
