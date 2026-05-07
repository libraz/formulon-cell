import { syncLayoutToEngine } from '../engine/layout-sync.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type {
  CellFormat,
  LayoutSlice,
  PageSetup,
  SlicerSpec,
  Sparkline,
  SpreadsheetStore,
  State,
} from '../store/store.js';

const LIMIT = 200;

/** A reversible operation. Both functions must be idempotent w.r.t. each other —
 *  calling `undo` then `redo` must leave the system in the same state. */
export interface HistoryEntry {
  undo: () => void;
  redo: () => void;
}

/**
 * Single source of truth for undoable mutations. Cell writes (workbook),
 * format changes, and layout changes (col widths, row heights, freeze) all
 * push entries here so one Cmd/Ctrl+Z spans the whole instance.
 *
 * Use `begin()` / `end()` to batch multiple entries into one logical step
 * (e.g. paste-special, fill drag).
 */
export class History {
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private replaying = false;
  private txnDepth = 0;
  private txnEntries: HistoryEntry[] = [];
  private listeners = new Set<() => void>();

  push(entry: HistoryEntry): void {
    if (this.replaying) return;
    if (this.txnDepth > 0) {
      this.txnEntries.push(entry);
      return;
    }
    this.commit(entry);
  }

  begin(): void {
    this.txnDepth += 1;
    if (this.txnDepth === 1) this.txnEntries = [];
  }

  end(): void {
    if (this.txnDepth === 0) return;
    this.txnDepth -= 1;
    if (this.txnDepth > 0) return;
    const entries = this.txnEntries;
    this.txnEntries = [];
    if (entries.length === 0) return;
    if (entries.length === 1) {
      const only = entries[0];
      if (only) this.commit(only);
      return;
    }
    this.commit({
      undo: () => {
        for (let i = entries.length - 1; i >= 0; i -= 1) entries[i]?.undo();
      },
      redo: () => {
        for (const e of entries) e.redo();
      },
    });
  }

  private commit(entry: HistoryEntry): void {
    this.undoStack.push(entry);
    if (this.undoStack.length > LIMIT) this.undoStack.shift();
    this.redoStack.length = 0;
    this.notify();
  }

  isReplaying(): boolean {
    return this.replaying;
  }

  undo(): boolean {
    const e = this.undoStack.pop();
    if (!e) return false;
    this.replaying = true;
    try {
      e.undo();
    } finally {
      this.replaying = false;
    }
    this.redoStack.push(e);
    this.notify();
    return true;
  }

  redo(): boolean {
    const e = this.redoStack.pop();
    if (!e) return false;
    this.replaying = true;
    try {
      e.redo();
    } finally {
      this.replaying = false;
    }
    this.undoStack.push(e);
    this.notify();
    return true;
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.txnEntries.length = 0;
    this.txnDepth = 0;
    this.notify();
  }

  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  private notify(): void {
    for (const l of this.listeners) l();
  }
}

/* ---------- Snapshot helpers ---------- */

/** Capture the entire format map. Sufficient for undo since the map is sparse
 *  (only explicitly formatted cells have entries). */
export function captureFormatSnapshot(state: State): Map<string, CellFormat> {
  return new Map(state.format.formats);
}

export function applyFormatSnapshot(store: SpreadsheetStore, snap: Map<string, CellFormat>): void {
  store.setState((s) => ({ ...s, format: { formats: new Map(snap) } }));
}

export interface LayoutSnapshot {
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  freezeRows: number;
  freezeCols: number;
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  outlineRows: Map<number, number>;
  outlineCols: Map<number, number>;
  outlineRowGutter: number;
  outlineColGutter: number;
  hiddenSheets: Set<number>;
}

export function captureLayoutSnapshot(state: State): LayoutSnapshot {
  return {
    colWidths: new Map(state.layout.colWidths),
    rowHeights: new Map(state.layout.rowHeights),
    freezeRows: state.layout.freezeRows,
    freezeCols: state.layout.freezeCols,
    hiddenRows: new Set(state.layout.hiddenRows),
    hiddenCols: new Set(state.layout.hiddenCols),
    outlineRows: new Map(state.layout.outlineRows),
    outlineCols: new Map(state.layout.outlineCols),
    outlineRowGutter: state.layout.outlineRowGutter,
    outlineColGutter: state.layout.outlineColGutter,
    hiddenSheets: new Set(state.layout.hiddenSheets),
  };
}

export function applyLayoutSnapshot(store: SpreadsheetStore, snap: LayoutSnapshot): void {
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      colWidths: new Map(snap.colWidths),
      rowHeights: new Map(snap.rowHeights),
      freezeRows: snap.freezeRows,
      freezeCols: snap.freezeCols,
      hiddenRows: new Set(snap.hiddenRows),
      hiddenCols: new Set(snap.hiddenCols),
      outlineRows: new Map(snap.outlineRows),
      outlineCols: new Map(snap.outlineCols),
      outlineRowGutter: snap.outlineRowGutter,
      outlineColGutter: snap.outlineColGutter,
      hiddenSheets: new Set(snap.hiddenSheets),
    } as LayoutSlice,
  }));
}

/** Run `mutate`, capturing the format slice before and after, pushing one
 *  entry. No-op when `history` is null. */
export function recordFormatChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureFormatSnapshot(store.getState());
  mutate();
  const after = captureFormatSnapshot(store.getState());
  history.push({
    undo: () => applyFormatSnapshot(store, before),
    redo: () => applyFormatSnapshot(store, after),
  });
}

export function recordLayoutChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureLayoutSnapshot(store.getState());
  mutate();
  const after = captureLayoutSnapshot(store.getState());
  history.push({
    undo: () => applyLayoutSnapshot(store, before),
    redo: () => applyLayoutSnapshot(store, after),
  });
}

/** Engine-aware layout change. Same semantics as `recordLayoutChange` but
 *  the captured before/after pair is also pushed to the workbook engine for
 *  the active sheet, both at apply time and on every undo/redo replay.
 *  Skipped (including the engine sync) when `wb` is null. The per-method
 *  calls inside `syncLayoutToEngine` short-circuit on each capability flag,
 *  so engines that only support a subset still work. */
export function recordLayoutChangeWithEngine(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  mutate: () => void,
): void {
  if (!wb) {
    recordLayoutChange(history, store, mutate);
    return;
  }
  const sheet = store.getState().data.sheetIndex;
  const before = captureLayoutSnapshot(store.getState());
  mutate();
  const after = captureLayoutSnapshot(store.getState());
  syncLayoutToEngine(wb, store.getState().layout, sheet, before, after);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => {
      applyLayoutSnapshot(store, before);
      syncLayoutToEngine(wb, store.getState().layout, sheet, after, before);
    },
    redo: () => {
      applyLayoutSnapshot(store, after);
      syncLayoutToEngine(wb, store.getState().layout, sheet, before, after);
    },
  });
}

export interface MergesSnapshot {
  byAnchor: Map<string, Range>;
  byCell: Map<string, string>;
}

export function captureMergesSnapshot(state: State): MergesSnapshot {
  return {
    byAnchor: new Map(state.merges.byAnchor),
    byCell: new Map(state.merges.byCell),
  };
}

export function applyMergesSnapshot(store: SpreadsheetStore, snap: MergesSnapshot): void {
  store.setState((s) => ({
    ...s,
    merges: { byAnchor: new Map(snap.byAnchor), byCell: new Map(snap.byCell) },
  }));
}

export function recordMergesChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureMergesSnapshot(store.getState());
  mutate();
  const after = captureMergesSnapshot(store.getState());
  history.push({
    undo: () => applyMergesSnapshot(store, before),
    redo: () => applyMergesSnapshot(store, after),
  });
}

/** Engine-aware merges change. Mirrors the post-mutate state into the workbook
 *  via `clearMerges` + per-anchor `addMerge`. Both apply and undo/redo go
 *  through the same path, so the engine snapshot stays in lockstep with the
 *  store. No-op against the engine when `wb` is null or `capabilities.merges`
 *  is off. */
export function recordMergesChangeWithEngine(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  sheet: number,
  mutate: () => void,
): void {
  const sync = (snap: MergesSnapshot): void => {
    if (!wb?.capabilities.merges) return;
    wb.engineClearMerges(sheet);
    for (const r of snap.byAnchor.values()) {
      if (r.sheet !== sheet) continue;
      wb.engineAddMerge(sheet, r);
    }
  };
  const before = captureMergesSnapshot(store.getState());
  mutate();
  const after = captureMergesSnapshot(store.getState());
  sync(after);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => {
      applyMergesSnapshot(store, before);
      sync(before);
    },
    redo: () => {
      applyMergesSnapshot(store, after);
      sync(after);
    },
  });
}

export function captureSparklineSnapshot(state: State): Map<string, Sparkline> {
  return new Map(state.sparkline.sparklines);
}

export function applySparklineSnapshot(
  store: SpreadsheetStore,
  snap: Map<string, Sparkline>,
): void {
  store.setState((s) => ({ ...s, sparkline: { sparklines: new Map(snap) } }));
}

export function recordSparklineChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureSparklineSnapshot(store.getState());
  mutate();
  const after = captureSparklineSnapshot(store.getState());
  history.push({
    undo: () => applySparklineSnapshot(store, before),
    redo: () => applySparklineSnapshot(store, after),
  });
}

export function capturePageSetupSnapshot(state: State): Map<number, PageSetup> {
  // Deep-clone each entry — margins is the only nested object so shallow copy
  // suffices but a fresh object guards against future field additions that
  // might re-use the same reference.
  const out = new Map<number, PageSetup>();
  for (const [k, v] of state.pageSetup.setupBySheet) {
    out.set(k, { ...v, margins: { ...v.margins } });
  }
  return out;
}

export function applyPageSetupSnapshot(
  store: SpreadsheetStore,
  snap: Map<number, PageSetup>,
): void {
  const next = new Map<number, PageSetup>();
  for (const [k, v] of snap) {
    next.set(k, { ...v, margins: { ...v.margins } });
  }
  store.setState((s) => ({ ...s, pageSetup: { setupBySheet: next } }));
}

/** Run `mutate`, capturing the page-setup slice before and after, pushing one
 *  entry. No-op when `history` is null. */
export function recordPageSetupChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = capturePageSetupSnapshot(store.getState());
  mutate();
  const after = capturePageSetupSnapshot(store.getState());
  history.push({
    undo: () => applyPageSetupSnapshot(store, before),
    redo: () => applyPageSetupSnapshot(store, after),
  });
}

/** Capture a deep-cloned slicer list for undo replay. Each spec is freshly
 *  cloned so future mutators that mutate the `selected` array can't
 *  retroactively pollute a prior snapshot. */
export function captureSlicersSnapshot(state: State): SlicerSpec[] {
  return state.slicers.slicers.map((sp) => ({ ...sp, selected: [...sp.selected] }));
}

export function applySlicersSnapshot(store: SpreadsheetStore, snap: readonly SlicerSpec[]): void {
  store.setState((s) => ({
    ...s,
    slicers: { slicers: snap.map((sp) => ({ ...sp, selected: [...sp.selected] })) },
  }));
}

/** Run `mutate` and push one history entry capturing the slicer-slice
 *  before/after. Use for any add/remove/update/setSelected call that should
 *  be undoable. */
export function recordSlicersChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureSlicersSnapshot(store.getState());
  mutate();
  const after = captureSlicersSnapshot(store.getState());
  history.push({
    undo: () => applySlicersSnapshot(store, before),
    redo: () => applySlicersSnapshot(store, after),
  });
}

/* ---------- Public undo/redo API ---------- */

/** Pull one entry off the undo stack and apply it. */
export function undo(history: History): boolean {
  return history.undo();
}

export function redo(history: History): boolean {
  return history.redo();
}

export function canUndo(history: History): boolean {
  return history.canUndo();
}

export function canRedo(history: History): boolean {
  return history.canRedo();
}

/** Legacy overloads — accept a `WorkbookHandle` so callers from before the
 *  unified history land keep working. The wb's internal stack is bypassed
 *  when a History is attached (mount does this by default). */
export function undoLegacy(wb: WorkbookHandle): boolean {
  return wb.undo();
}

export function redoLegacy(wb: WorkbookHandle): boolean {
  return wb.redo();
}
