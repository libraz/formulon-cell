import type { Addr, Range } from '../engine/types.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';

const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

export function isWatched(store: SpreadsheetStore, addr: Addr): boolean {
  return store.getState().watch.watches.some((w) => sameAddr(w, addr));
}

export function watchCell(store: SpreadsheetStore, addr: Addr): void {
  mutators.addWatch(store, addr);
}

export function watchRange(store: SpreadsheetStore, range: Range): void {
  mutators.addWatchRange(store, range);
}

export function watchRanges(store: SpreadsheetStore, ranges: readonly Range[]): void {
  mutators.addWatchRanges(store, ranges);
}

export function recordWatchesChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => T,
): T {
  const before = store.getState().watch.watches.map((addr) => ({ ...addr }));
  const result = mutate();
  const after = store.getState().watch.watches.map((addr) => ({ ...addr }));
  const same =
    before.length === after.length &&
    before.every((addr, index) => {
      const other = after[index];
      return (
        other !== undefined &&
        addr.sheet === other.sheet &&
        addr.row === other.row &&
        addr.col === other.col
      );
    });
  if (history && !history.isReplaying() && !same) {
    history.push({
      undo: () => mutators.setWatches(store, before),
      redo: () => mutators.setWatches(store, after),
    });
  }
  return result;
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
