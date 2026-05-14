import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

export function isCellErrorIgnored(store: SpreadsheetStore, addr: Addr): boolean {
  return store.getState().errorIndicators.ignoredErrors.has(addrKey(addr));
}

export function ignoreCellError(store: SpreadsheetStore, addr: Addr): void {
  mutators.ignoreError(store, addr);
}

export function restoreCellErrorIndicator(store: SpreadsheetStore, addr: Addr): void {
  mutators.unignoreError(store, addr);
}

export function toggleCellErrorIgnored(store: SpreadsheetStore, addr: Addr): boolean {
  if (isCellErrorIgnored(store, addr)) {
    restoreCellErrorIndicator(store, addr);
    return false;
  }
  ignoreCellError(store, addr);
  return true;
}

export function clearIgnoredCellErrors(store: SpreadsheetStore): void {
  mutators.clearIgnoredErrors(store);
}
