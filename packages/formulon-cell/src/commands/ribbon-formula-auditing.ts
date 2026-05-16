// Shared "Formula Auditing" ribbon split-button. The host wrappers used to
// duplicate this branch-by-branch — now they hand us the action and we either:
//   - return `{kind:'trace-precedents'}` so the host can invoke its
//     viewport-aware tracePrecedents method, or
//   - run the matching error-indicator / validation-circle mutation here, then
//     advance the selection to the next formula error and surface an
//     "all clear" report when nothing remains.
//
// Branches that just need a no-match report still come back as
// `{kind:'report', report}` so each host renders through its own ribbon-report
// shell.

import { makeRangeResolver } from '../engine/range-resolver.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import {
  cellValueIsFormulaError,
  circleInvalidValidationData,
  clearValidationCircles,
  ignoreCellError,
  recordIgnoredErrorsChange,
  recordValidationCirclesChange,
  selectNextFormulaError,
} from './error-indicators.js';
import type { History } from './history.js';

type FormulaAuditingStrings = Pick<Strings['ribbonMenu'], 'errorChecking'>;

export type RibbonFormulaAuditingAction =
  | 'errorChecking'
  | 'traceError'
  | 'ignoreError'
  | 'circleInvalid'
  | 'clearCircles';

export interface RibbonFormulaAuditingReport {
  title: string;
  items: { severity: 'info' | 'warning'; label: string; detail: string }[];
}

export type RibbonFormulaAuditingActionResult =
  | { kind: 'trace-precedents' }
  | { kind: 'mutated' }
  | { kind: 'report'; report: RibbonFormulaAuditingReport };

export interface ExecuteRibbonFormulaAuditingActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  action: RibbonFormulaAuditingAction;
  strings: FormulaAuditingStrings;
}

const noErrorReport = (strings: FormulaAuditingStrings): RibbonFormulaAuditingReport => ({
  title: strings.errorChecking,
  items: [],
});

/** Dispatch one formula-auditing action. Branches:
 *  - `traceError` → host runs `tracePrecedents` (visual arrow drawing).
 *  - `ignoreError` → on a formula-error cell, mark ignored; otherwise advance
 *    to the next formula error and report "all clear" when none remain.
 *  - `circleInvalid` → draw invalid-data-validation circles for the active
 *    range, scoped via [[makeRangeResolver]].
 *  - `clearCircles` → wipe every validation circle on the sheet.
 *  - `errorChecking` (default) → advance to the next formula error or report. */
export const executeRibbonFormulaAuditingAction = (
  deps: ExecuteRibbonFormulaAuditingActionDeps,
): RibbonFormulaAuditingActionResult => {
  const { store, workbook, history, action, strings } = deps;
  if (action === 'traceError') return { kind: 'trace-precedents' };
  if (action === 'ignoreError') {
    const state = store.getState();
    const active = state.selection.active;
    const activeCell = state.data.cells.get(`${active.sheet}:${active.row}:${active.col}`);
    if (activeCell?.formula && cellValueIsFormulaError(activeCell.value)) {
      recordIgnoredErrorsChange(history, store, () => {
        ignoreCellError(store, active);
      });
      return { kind: 'mutated' };
    }
    return selectNextFormulaError(store)
      ? { kind: 'mutated' }
      : { kind: 'report', report: noErrorReport(strings) };
  }
  if (action === 'clearCircles') {
    recordValidationCirclesChange(history, store, () => {
      clearValidationCircles(store);
    });
    return { kind: 'mutated' };
  }
  if (action === 'circleInvalid') {
    const state = store.getState();
    recordValidationCirclesChange(history, store, () => {
      circleInvalidValidationData(
        store,
        state.selection.range,
        makeRangeResolver(workbook, state.data.sheetIndex),
      );
    });
    return { kind: 'mutated' };
  }
  return selectNextFormulaError(store)
    ? { kind: 'mutated' }
    : { kind: 'report', report: noErrorReport(strings) };
};
