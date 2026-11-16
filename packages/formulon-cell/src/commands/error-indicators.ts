import { addrKey, parseAddrKey } from '../engine/address.js';
import { calcModeLeavesStaleValues } from '../engine/calc-mode.js';
import { effectiveCellFormat } from '../engine/cell-format-projection.js';
import type { RangeResolver } from '../engine/range-resolver.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { ERROR_LITERALS } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeValidationCells } from '../store/range-validations.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import {
  type CellValidation,
  type ErrorCheckingSettings,
  type ErrorCheckRuleId,
  mutators,
  type SpreadsheetStore,
} from '../store/store.js';
import { commitCellInput } from './cell-input-commit.js';
import { inputLocaleForWorkbook, parseLocalizedNumericInput } from './coerce-input.js';
import type { History } from './history.js';
import { formatA1FormulaAsR1C1 } from './refs.js';
import { cellValueViolatesValidation, circlesOutsideScope } from './validate.js';

export function cellValueIsFormulaError(value: CellValue): boolean {
  if (value.kind === 'error') return true;
  if (value.kind === 'text') return ERROR_LITERALS.has(value.value);
  return false;
}

/** The spreadsheet's "formulas that result in an error" check, which puts
 *  the indicator on a cell and drives Error Checking: a formula whose value
 *  is an error. A bare `=NA()` marks a value as missing on purpose and is not
 *  flagged, nor is a text result that only reads like an error. The Mac
 *  desktop (`macNa`) leaves every #N/A unflagged, whatever formula gave it. */
export function formulaErrorIsFlagged(
  value: CellValue,
  formula: string | null | undefined,
  opts: { macNa?: boolean } = {},
): boolean {
  if (!formula || value.kind !== 'error') return false;
  if (opts.macNa && value.code === 6) return false;
  return formula.replace(/^=/, '').replace(/\s+/g, '').toUpperCase() !== 'NA()';
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

const rowMajor = (left: Addr, right: Addr): number =>
  left.row - right.row || left.col - right.col || left.sheet - right.sheet;

/** What Error Checking needs to flag the cells the grid marks: the profile's
 *  #N/A rule and number-stored-as-text check, the input locale numbers are
 *  read in, and the workbook whose blocked PivotTables show #SPILL!. */
export interface FormulaErrorCheckOptions {
  macNa?: boolean;
  numberAsText?: boolean;
  /** The formula rules of the desktop's default check: a formula inconsistent
   *  with its neighbours, one that omits adjacent data, an unlocked formula. */
  formulaRules?: boolean;
  locale?: string;
  /** The Preferences › Error Checking flags; absent means every rule is on. */
  errorChecking?: ErrorCheckingSettings;
  workbook?: Pick<WorkbookHandle, 'pivotSpills'> | null;
}

/** A neighbour's content, for the rules that compare a cell with its region. */
export type ErrorCheckCellAt = (
  row: number,
  col: number,
) => { value: CellValue; formula: string | null } | undefined;

/** The cell being checked, as far as the formula rules need it. */
export interface ErrorCheckCell {
  addr: Addr;
  locked?: boolean;
  cellAt: ErrorCheckCellAt;
}

const relativeForm = (formula: string, addr: Addr): string =>
  formatA1FormulaAsR1C1(formula, addr).replace(/\s+/g, '').toUpperCase();

/** The formula differs from neighbours on both sides that agree with each
 *  other, down its column or across its row. */
function inconsistentWithRegion(formula: string, cell: ErrorCheckCell): boolean {
  const { row, col } = cell.addr;
  const own = relativeForm(formula, cell.addr);
  const formOf = (r: number, c: number): string | null => {
    const f = r >= 0 && c >= 0 ? cell.cellAt(r, c)?.formula : null;
    return f ? relativeForm(f, { sheet: cell.addr.sheet, row: r, col: c }) : null;
  };
  for (const [a, b] of [
    [formOf(row - 1, col), formOf(row + 1, col)],
    [formOf(row, col - 1), formOf(row, col + 1)],
  ] as const) {
    if (a !== null && a === b && a !== own) return true;
  }
  return false;
}

const isNumber = (cell: ReturnType<ErrorCheckCellAt>): boolean =>
  cell?.value.kind === 'number' && !cell.formula;

/** A range the formula refers to stops short of numbers between it and the
 *  formula, down the formula's column or across its row. */
function omitsAdjacentCells(formula: string, cell: ErrorCheckCell): boolean {
  const { row, col } = cell.addr;
  for (const match of formula.matchAll(/(?<![!\w$])\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+/gi)) {
    const ref = parseRangeRef(match[0]);
    if (!ref) continue;
    if (
      ref.c0 === col &&
      ref.c1 === col &&
      ref.r1 < row - 1 &&
      isNumber(cell.cellAt(ref.r1 + 1, col))
    ) {
      return true;
    }
    if (
      ref.r0 === row &&
      ref.r1 === row &&
      ref.c1 < col - 1 &&
      isNumber(cell.cellAt(row, ref.c1 + 1))
    ) {
      return true;
    }
  }
  return false;
}

/** The formula rules of the default error check, for a formula whose value is
 *  not an error. */
function formulaRuleFlagged(
  formula: string,
  cell: ErrorCheckCell,
  on: (rule: ErrorCheckRuleId) => boolean,
): boolean {
  return (
    (on('unlocked') && cell.locked === false) ||
    (on('inconsistent') && inconsistentWithRegion(formula, cell)) ||
    (on('omitsCells') && omitsAdjacentCells(formula, cell))
  );
}

/** A text constant that typing would have entered as a number: the desktop's
 *  "number stored as text" check. Date-like text such as "12/3" is not one. */
export function numberStoredAsText(
  value: CellValue,
  formula: string | null | undefined,
  locale?: string,
): boolean {
  if (formula || value.kind !== 'text') return false;
  return parseLocalizedNumericInput(value.value, locale) !== null;
}

/** The Convert to Number action of the number-stored-as-text rule: the text
 *  becomes the number it reads as, and a Text format gives way to General. */
export function convertTextToNumber(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  addr: Addr,
): boolean {
  const state = store.getState();
  const cell = state.data.cells.get(addrKey(addr));
  if (cell?.value.kind !== 'text') return false;
  if (parseLocalizedNumericInput(cell.value.value, inputLocaleForWorkbook(wb)) === null) {
    return false;
  }
  if (effectiveCellFormat(wb, state, addr)?.numFmt?.kind === 'text') {
    mutators.setCellFormat(store, addr, { numFmt: { kind: 'general' } });
  }
  const result = commitCellInput({
    store,
    wb,
    addr,
    raw: cell.value.value.trim(),
    origin: 'formulaBar',
  });
  return result.status === 'applied';
}

/** Whether stale formula values are marked: Format Stale Values is on, the calc mode can leave
 *  values stale, and the engine reports staleness. */
export const staleValuesMarked = (
  wb: Pick<WorkbookHandle, 'calcMode' | 'staleValuesSupported'> | null | undefined,
  formatStaleValues: boolean,
): boolean =>
  !!wb &&
  formatStaleValues &&
  calcModeLeavesStaleValues(wb.calcMode?.() ?? null) &&
  wb.staleValuesSupported?.() === true;

/** Whether the formula cell at `addr` holds a stale value the grid marks. */
export const staleFormulaMarked = (
  wb: Pick<WorkbookHandle, 'calcMode' | 'staleValuesSupported' | 'isCellStale'> | null | undefined,
  formatStaleValues: boolean,
  formula: string | null | undefined,
  addr: Addr,
): boolean => !!formula && staleValuesMarked(wb, formatStaleValues) && !!wb?.isCellStale(addr);

/** Whether a cell carries the error-check triangle: the one rule the grid
 *  paints by and Error Checking walks. */
export function errorCheckFlagged(
  value: CellValue,
  formula: string | null | undefined,
  pivotSpill: boolean,
  opts: Pick<
    FormulaErrorCheckOptions,
    'macNa' | 'numberAsText' | 'formulaRules' | 'locale' | 'errorChecking'
  > = {},
  cell?: ErrorCheckCell,
): boolean {
  if (opts.errorChecking?.enabled === false) return false;
  const on = (rule: ErrorCheckRuleId): boolean => opts.errorChecking?.rules[rule] !== false;
  return (
    (on('error') && (pivotSpill || formulaErrorIsFlagged(value, formula, opts))) ||
    (opts.numberAsText === true &&
      on('numberAsText') &&
      numberStoredAsText(value, formula, opts.locale)) ||
    (opts.formulaRules === true &&
      !!formula &&
      value.kind !== 'error' &&
      cell !== undefined &&
      formulaRuleFlagged(formula, cell, on))
  );
}

export function formulaErrorCellsInRange(
  store: SpreadsheetStore,
  range?: Range,
  opts: FormulaErrorCheckOptions = {},
): Addr[] {
  const state = store.getState();
  const target = range ?? state.selection.range;
  const anchors = new Set(
    (typeof opts.workbook?.pivotSpills === 'function'
      ? opts.workbook.pivotSpills(state.data.sheetIndex)
      : []
    ).map((spill) => addrKey(spill.anchor)),
  );
  const out: Addr[] = [];
  const cellAt: ErrorCheckCellAt = (row, col) =>
    state.data.cells.get(addrKey({ sheet: state.data.sheetIndex, row, col }));
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr) continue;
    const locked = state.format.formats.get(key)?.locked;
    if (
      !errorCheckFlagged(cell.value, cell.formula, anchors.has(key), opts, {
        addr,
        ...(locked !== undefined ? { locked } : {}),
        cellAt,
      })
    ) {
      continue;
    }
    if (state.errorIndicators.ignoredErrors.has(key)) continue;
    if (rangeContainsAddr(target, addr)) out.push(addr);
  }
  return out.sort(rowMajor);
}

export function selectNextFormulaError(
  store: SpreadsheetStore,
  range?: Range,
  accept?: (addr: Addr) => boolean,
  opts: FormulaErrorCheckOptions = {},
): Addr | null {
  const state = store.getState();
  const active = state.selection.active;
  const errors = formulaErrorCellsInRange(store, range ?? state.selection.range, opts).filter(
    (addr) => accept?.(addr) ?? true,
  );
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
  const keys = circlesOutsideScope(state.errorIndicators.validationCircles, range.sheet, range);
  let marked = 0;
  const candidates: Array<{ key: string; validation: CellValidation }> = [];
  for (const [key, format] of state.format.formats) {
    if (!format.validation) continue;
    const addr = parseAddrKey(key);
    if (!addr || !rangeContainsAddr(range, addr)) continue;
    candidates.push({ key, validation: format.validation });
  }
  for (const { addr, validation } of rangeValidationCells(state, range.sheet)) {
    if (rangeContainsAddr(range, addr)) candidates.push({ key: addrKey(addr), validation });
  }
  candidates.sort((left, right) => {
    const leftAddr = parseAddrKey(left.key);
    const rightAddr = parseAddrKey(right.key);
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
  const keys = circlesOutsideScope(state.errorIndicators.validationCircles, sheet);
  let marked = 0;
  for (const [key, format] of state.format.formats) {
    if (!format.validation) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const value = state.data.cells.get(key)?.value ?? { kind: 'blank' as const };
    if (!cellValueViolatesValidation(value, format.validation, resolveRange)) continue;
    keys.add(key);
    marked += 1;
  }
  for (const { addr, validation } of rangeValidationCells(state, sheet)) {
    const key = addrKey(addr);
    const value = state.data.cells.get(key)?.value ?? { kind: 'blank' as const };
    if (!cellValueViolatesValidation(value, validation, resolveRange)) continue;
    keys.add(key);
    marked += 1;
  }
  mutators.setValidationCircles(store, keys);
  return marked;
}

export function clearValidationCircles(store: SpreadsheetStore): void {
  mutators.clearValidationCircles(store);
}
