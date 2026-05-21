import { findDependents, findPrecedents } from '../engine/refs-graph.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type TraceArrow } from '../store/store.js';
import type { History } from './history.js';

const cloneAddr = (addr: TraceArrow['from']): TraceArrow['from'] => ({ ...addr });

const cloneTrace = (arrow: TraceArrow): TraceArrow => ({
  kind: arrow.kind,
  from: cloneAddr(arrow.from),
  to: cloneAddr(arrow.to),
});

const captureTraceSnapshot = (store: SpreadsheetStore): TraceArrow[] =>
  store.getState().traces.items.map(cloneTrace);

const applyTraceSnapshot = (store: SpreadsheetStore, snap: readonly TraceArrow[]): void => {
  store.setState((s) => ({ ...s, traces: { items: snap.map(cloneTrace) } }));
};

const sameAddr = (a: TraceArrow['from'], b: TraceArrow['from']): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

const sameTrace = (a: TraceArrow, b: TraceArrow): boolean =>
  a.kind === b.kind && sameAddr(a.from, b.from) && sameAddr(a.to, b.to);

const sameTraceSnapshot = (a: readonly TraceArrow[], b: readonly TraceArrow[]): boolean =>
  a.length === b.length &&
  a.every((arrow, index) => {
    const other = b[index];
    return other ? sameTrace(arrow, other) : false;
  });

export function recordTraceChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => T,
): T {
  if (!history || history.isReplaying()) return mutate();
  const before = captureTraceSnapshot(store);
  const result = mutate();
  const after = captureTraceSnapshot(store);
  if (!sameTraceSnapshot(before, after)) {
    history.push({
      undo: () => applyTraceSnapshot(store, before),
      redo: () => applyTraceSnapshot(store, after),
    });
  }
  return result;
}

export function addTraceArrow(store: SpreadsheetStore, arrow: TraceArrow): void {
  mutators.addTrace(store, arrow);
}

export function tracePrecedents(
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  addr = store.getState().selection.active,
  history: History | null = null,
): number {
  return recordTraceChange(history, store, () => {
    const precedents = findPrecedents(workbook, addr);
    for (const from of precedents) {
      addTraceArrow(store, { kind: 'precedent', from, to: addr });
    }
    return precedents.length;
  });
}

export function traceDependents(
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  addr = store.getState().selection.active,
  history: History | null = null,
): number {
  return recordTraceChange(history, store, () => {
    const dependents = findDependents(workbook, addr);
    for (const to of dependents) {
      addTraceArrow(store, { kind: 'dependent', from: addr, to });
    }
    return dependents.length;
  });
}

export function clearTraceArrows(store: SpreadsheetStore, history: History | null = null): void {
  recordTraceChange(history, store, () => {
    mutators.clearTraces(store);
  });
}

export function clearTraceArrowsByKind(
  store: SpreadsheetStore,
  kind: TraceArrow['kind'],
  history: History | null = null,
): void {
  recordTraceChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      traces: { items: s.traces.items.filter((arrow) => arrow.kind !== kind) },
    }));
  });
}
