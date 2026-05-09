import type { Addr } from '../engine/types.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

export function isWatched(store: SpreadsheetStore, addr: Addr): boolean {
  return store.getState().watch.watches.some((w) => sameAddr(w, addr));
}

export function watchCell(store: SpreadsheetStore, addr: Addr): void {
  mutators.addWatch(store, addr);
}

export function unwatchCell(store: SpreadsheetStore, addr: Addr): void {
  mutators.removeWatch(store, addr);
}

export function toggleWatchCell(store: SpreadsheetStore, addr: Addr): boolean {
  if (isWatched(store, addr)) {
    unwatchCell(store, addr);
    return false;
  }
  watchCell(store, addr);
  return true;
}

export function clearWatchedCells(store: SpreadsheetStore): void {
  mutators.clearWatches(store);
}

export function setWatchWindowOpen(store: SpreadsheetStore, open: boolean): void {
  mutators.setWatchPanelOpen(store, open);
}
