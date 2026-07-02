import { addrKey } from '../engine/address.js';
import type { RangeResolver } from '../engine/range-resolver.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { type CellValidation, mutators, type SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { cellValueViolatesValidation } from './validate.js';

const ERROR_SENTINELS: ReadonlySet<string> = new Set([
  '#DIV/0!',
  '#NAME?',
  '#REF!',
  '#VALUE!',
  '#NUM!',
  '#N/A',
  '#NULL!',
  '#CIRCULAR!',
]);

export function cellValueIsFormulaError(value: CellValue): boolean {
  if (value.kind === 'error') return true;
  if (value.kind === 'text') return ERROR_SENTINELS.has(value.value);
  return false;
}

export function recordValidationCirclesChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => T,
): T {
  const before = new Set(store.getState().errorIndicators.validationCircles);
  const result = mutate();
  const after = new Set(store.getState().errorIndicators.validationCircles);
  const same = before.size === after.size && Array.from(before).every((key) => after.has(key));
  if (history && !history.isReplaying() && !same) {
    history.push({
      undo: () => mutators.setValidationCircles(store, before),
      redo: () => mutators.setValidationCircles(store, after),
    });
  }
  return result;
}

export function recordIgnoredErrorsChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => T,
): T {
  const before = new Set(store.getState().errorIndicators.ignoredErrors);
  const result = mutate();
  const after = new Set(store.getState().errorIndicators.ignoredErrors);
  const same = before.size === after.size && Array.from(before).every((key) => after.has(key));
  if (history && !history.isReplaying() && !same) {
    history.push({
      undo: () => mutators.setIgnoredErrors(store, before),
      redo: () => mutators.setIgnoredErrors(store, after),
    });
  }
  return result;
}

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

const rangeContainsAddr = (range: Range, addr: Addr): boolean =>
  addr.sheet === range.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

const addrFromKey = (key: string): Addr | null => {
  const parts = key.split(':').map(Number);
  const sheet = parts[0];
  const row = parts[1];
  const col = parts[2];
  if (
    typeof sheet !== 'number' ||
    typeof row !== 'number' ||
    typeof col !== 'number' ||
    !Number.isInteger(sheet) ||
    !Number.isInteger(row) ||
    !Number.isInteger(col)
  ) {
    return null;
  }
  return { sheet, row, col };
};

const rowMajor = (left: Addr, right: Addr): number =>
  left.row - right.row || left.col - right.col || left.sheet - right.sheet;

export function formulaErrorCellsInRange(store: SpreadsheetStore, range?: Range): Addr[] {
  const state = store.getState();
  const target = range ?? state.selection.range;
  const out: Addr[] = [];
  for (const [key, cell] of state.data.cells) {
    if (!cell.formula || !cellValueIsFormulaError(cell.value)) continue;
    if (state.errorIndicators.ignoredErrors.has(key)) continue;
    const addr = addrFromKey(key);
    if (addr && rangeContainsAddr(target, addr)) out.push(addr);
  }
  return out.sort(rowMajor);
}

export function selectNextFormulaError(store: SpreadsheetStore, range?: Range): Addr | null {
  const state = store.getState();
  const active = state.selection.active;
  const errors = formulaErrorCellsInRange(store, range ?? state.selection.range);
  if (errors.length === 0) return null;
  const afterActive =
    errors.find(
      (addr) => addr.row > active.row || (addr.row === active.row && addr.col > active.col),
    ) ?? errors[0];
  if (!afterActive) return null;
  mutators.setActive(store, afterActive);
  return afterActive;
}

export function circleInvalidValidationData(
  store: SpreadsheetStore,
  range: Range,
  resolveRange?: RangeResolver,
): number {
  const state = store.getState();
  const keys = new Set(state.errorIndicators.validationCircles);
  let marked = 0;
  const candidates: Array<{ key: string; validation: CellValidation }> = [];
  for (const [key, format] of state.format.formats) {
    if (!format.validation) continue;
    const addr = addrFromKey(key);
    if (!addr || !rangeContainsAddr(range, addr)) continue;
    candidates.push({ key, validation: format.validation });
  }
  candidates.sort((left, right) => {
    const leftAddr = addrFromKey(left.key);
    const rightAddr = addrFromKey(right.key);
    return leftAddr && rightAddr ? rowMajor(leftAddr, rightAddr) : 0;
  });
  for (const { key, validation } of candidates) {
    const value = state.data.cells.get(key)?.value ?? { kind: 'blank' as const };
    if (!cellValueViolatesValidation(value, validation, resolveRange)) continue;
    keys.add(key);
    marked += 1;
  }
  mutators.setValidationCircles(store, keys);
  return marked;
}

export function circleInvalidValidationDataInSheet(
  store: SpreadsheetStore,
  sheet: number,
  resolveRange?: RangeResolver,
): number {
  const state = store.getState();
  const keys = new Set(state.errorIndicators.validationCircles);
  let marked = 0;
  for (const [key, format] of state.format.formats) {
    if (!format.validation) continue;
    const [keySheet, row, col] = key.split(':').map(Number);
    if (keySheet !== sheet || row === undefined || col === undefined) continue;
    const value = state.data.cells.get(key)?.value ?? { kind: 'blank' as const };
    if (!cellValueViolatesValidation(value, format.validation, resolveRange)) continue;
    keys.add(key);
    marked += 1;
  }
  mutators.setValidationCircles(store, keys);
  return marked;
}

export function clearValidationCircles(store: SpreadsheetStore): void {
  mutators.clearValidationCircles(store);
}
