import { findDependents, findPrecedents } from '../engine/refs-graph.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type TraceArrow } from '../store/store.js';

export function addTraceArrow(store: SpreadsheetStore, arrow: TraceArrow): void {
  mutators.addTrace(store, arrow);
}

export function tracePrecedents(
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  addr = store.getState().selection.active,
): number {
  const precedents = findPrecedents(workbook, addr);
  for (const from of precedents) {
    addTraceArrow(store, { kind: 'precedent', from, to: addr });
  }
  return precedents.length;
}

export function traceDependents(
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  addr = store.getState().selection.active,
): number {
  const dependents = findDependents(workbook, addr);
  for (const to of dependents) {
    addTraceArrow(store, { kind: 'dependent', from: addr, to });
  }
  return dependents.length;
}

export function clearTraceArrows(store: SpreadsheetStore): void {
  mutators.clearTraces(store);
}
