import { recordRecentFunction } from '../../commands/function-history.js';
import { addrKey } from '../../engine/address.js';
import {
  type AutoSumFunction,
  autoSum,
  cellValueIsFormulaError,
  clearTraceArrowsByKind,
  clearWatchedCells,
  createDefinedNamesFromSelection,
  executeRibbonFormulaAuditingAction,
  insertDefinedNameFormula,
  listDefinedNames,
  mutators,
  recordDefinedNamesChange,
  recordWatchesChange,
  type SpreadsheetInstance,
  unwatchCell,
  watchRange,
} from '../../index.js';
import { showDefinedNamePickerDialog } from '../../toolbar/dialogs/defined-name-picker.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';

const buildAutoSumFormula =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyAutoSumFormula'] =>
  (fn) => {
    if (fn === 'MORE') {
      instance.openFunctionArguments();
      return;
    }
    instance.history.begin();
    let result: ReturnType<typeof autoSum> = null;
    try {
      result = autoSum(instance.store.getState(), instance.workbook, fn as AutoSumFunction);
    } finally {
      instance.history.end();
    }
    if (result) {
      mutators.setActive(instance.store, result.addr);
      recordRecentFunction(instance.store, fn);
    }
    instance.host.focus();
  };

const buildWatchAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyWatchAction'] =>
  (action) => {
    const state = instance.store.getState();
    if (action === 'open') {
      instance.openWatchWindow();
      return;
    }
    if (action === 'add') {
      recordWatchesChange(instance.history, instance.store, () => {
        watchRange(instance.store, state.selection.range);
      });
      instance.openWatchWindow();
      return;
    }
    if (action === 'delete') {
      recordWatchesChange(instance.history, instance.store, () => {
        unwatchCell(instance.store, state.selection.active);
      });
      instance.openWatchWindow();
      return;
    }
    if (action === 'delete-all') {
      recordWatchesChange(instance.history, instance.store, () => {
        clearWatchedCells(instance.store);
      });
      instance.openWatchWindow();
    }
  };

const updateWatchMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateWatchMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const hasWatches = state.watch.watches.length > 0;
    const activeWatched = state.watch.watches.some(
      (watch) =>
        watch.sheet === active.sheet && watch.row === active.row && watch.col === active.col,
    );
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-watch-action]')) {
      const action = button.dataset.watchAction;
      const disabled =
        (action === 'delete' && !activeWatched) || (action === 'delete-all' && !hasWatches);
      const reason =
        action === 'delete' ? strings.watchDeleteRequiresActive : strings.watchDeleteAllRequiresAny;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildCalcOptionAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCalcOptionAction'] =>
  (action) => {
    if (action === 'auto' || action === 'manual' || action === 'auto-no-table') {
      const mode = action === 'auto' ? 0 : action === 'manual' ? 1 : 2;
      instance.workbook.setCalcMode(mode as 0 | 1 | 2);
      instance.host.focus();
      return;
    }
    if (action === 'calculate-now' || action === 'calculate-sheet') {
      instance.recalc();
      instance.host.focus();
      return;
    }
    if (action === 'iterative') {
      instance.openIterativeDialog();
    }
  };

const calcOptionForMode = (mode: 0 | 1 | 2 | null): 'auto' | 'manual' | 'auto-no-table' | null => {
  if (mode === 0) return 'auto';
  if (mode === 1) return 'manual';
  if (mode === 2) return 'auto-no-table';
  return null;
};

const updateCalcOptionsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCalcOptionsMenu'] =>
  (menu) => {
    const current = calcOptionForMode(instance.workbook.calcMode());
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]')) {
      const active = button.dataset.calcOption === current;
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('fc-tb__menu-item--active', active);
    }
  };

const buildFormulaAuditAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyFormulaAuditAction'] =>
  async (action) => {
    if (action === 'precedents') {
      instance.tracePrecedents();
      return;
    }
    if (action === 'dependents') {
      instance.traceDependents();
      return;
    }
    if (action === 'clear-all') {
      instance.clearTraces();
      return;
    }
    if (action === 'clear-precedents' || action === 'clear-dependents') {
      clearTraceArrowsByKind(
        instance.store,
        action === 'clear-precedents' ? 'precedent' : 'dependent',
        instance.history,
      );
      instance.host.focus();
      return;
    }
    const map: Record<string, 'errorChecking' | 'traceError' | 'ignoreError'> = {
      'error-checking': 'errorChecking',
      'trace-error': 'traceError',
      'ignore-error': 'ignoreError',
    };
    const auditAction = map[action];
    if (!auditAction) return;
    const result = executeRibbonFormulaAuditingAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action: auditAction,
      strings: { errorChecking: instance.i18n.strings.ribbonMenu.errorChecking },
    });
    if (result.kind === 'trace-precedents') {
      instance.tracePrecedents();
      return;
    }
    if (result.kind === 'report') {
      await showInstanceReport(instance, result.report.title, result.report.items);
    }
    instance.host.focus();
  };

const updateErrorCheckingMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateErrorCheckingMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const cell = state.data.cells.get(addrKey(active));
    const activeFormulaError = Boolean(cell?.formula && cellValueIsFormulaError(cell.value));
    const disabledReason = instance.i18n.strings.ribbonMenu.traceErrorRequiresFormulaError;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-formula-audit-action]')) {
      const action = button.dataset.formulaAuditAction;
      const disabled =
        (action === 'trace-error' || action === 'ignore-error') && !activeFormulaError;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const updateClearArrowsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateClearArrowsMenu'] =>
  (menu) => {
    const traces = instance.store.getState().traces.items;
    const hasPrecedents = traces.some((trace) => trace.kind === 'precedent');
    const hasDependents = traces.some((trace) => trace.kind === 'dependent');
    const hasAny = hasPrecedents || hasDependents;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-formula-audit-action]')) {
      const action = button.dataset.formulaAuditAction;
      const disabled =
        (action === 'clear-all' && !hasAny) ||
        (action === 'clear-precedents' && !hasPrecedents) ||
        (action === 'clear-dependents' && !hasDependents);
      const reason =
        action === 'clear-precedents'
          ? strings.removePrecedentArrowsRequiresAny
          : action === 'clear-dependents'
            ? strings.removeDependentArrowsRequiresAny
            : strings.removeArrowsRequiresAny;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildDefinedNameAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyDefinedNameAction'] =>
  async (action) => {
    if (action === 'define') {
      instance.openDefineNameDialog();
      return;
    }
    if (action === 'manager') {
      instance.openNamedRangeDialog();
      return;
    }
    const source =
      action === 'create-top-row'
        ? 'top-row'
        : action === 'create-bottom-row'
          ? 'bottom-row'
          : action === 'create-left-column'
            ? 'left-column'
            : action === 'create-right-column'
              ? 'right-column'
              : null;
    if (source) {
      recordDefinedNamesChange(instance.history, instance.workbook, () =>
        createDefinedNamesFromSelection(instance.store.getState(), instance.workbook, source),
      );
      instance.host.focus();
      return;
    }
    if (action === 'use-formula') {
      const names = listDefinedNames(instance.workbook);
      if (names.length === 0) {
        await showInstanceReport(instance, instance.i18n.strings.ribbonMenu.useInFormula, [
          {
            severity: 'info',
            label: instance.i18n.strings.ribbonMenu.noDefinedNames,
            detail: '',
          },
        ]);
        return;
      }
      const selected = await showDefinedNamePickerDialog({
        title: instance.i18n.strings.ribbonMenu.useInFormula,
        names,
        okLabel: instance.i18n.strings.hyperlinkDialog.ok,
        cancelLabel: instance.i18n.strings.hyperlinkDialog.cancel,
      });
      if (!selected) {
        instance.host.focus();
        return;
      }
      insertDefinedNameFormula(
        instance.store.getState(),
        instance.workbook,
        selected,
        instance.store,
      );
      mutators.replaceCells(
        instance.store,
        instance.workbook.cells(instance.store.getState().data.sheetIndex),
      );
      instance.host.focus();
    }
  };

const updateDefinedNamesMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateDefinedNamesMenu'] =>
  (menu) => {
    const hasDefinedNames = listDefinedNames(instance.workbook).length > 0;
    const canMutateNames = instance.workbook.capabilities.definedNameMutate;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-defined-name-action]')) {
      const action = button.dataset.definedNameAction;
      const disabled =
        action === 'use-formula'
          ? !hasDefinedNames
          : action?.startsWith('create-') === true && !canMutateNames;
      const reason =
        action === 'use-formula' ? strings.noDefinedNames : strings.definedNameMutationUnavailable;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

type FormulasDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyAutoSumFormula'
  | 'applyFormulaAuditAction'
  | 'updateErrorCheckingMenu'
  | 'updateClearArrowsMenu'
  | 'applyWatchAction'
  | 'updateWatchMenu'
  | 'applyCalcOptionAction'
  | 'updateCalcOptionsMenu'
  | 'applyDefinedNameAction'
  | 'updateDefinedNamesMenu'
>;

export function createFormulasDropdownDefaults(
  instance: SpreadsheetInstance,
): FormulasDropdownDefaults {
  return {
    applyAutoSumFormula: buildAutoSumFormula(instance),
    applyFormulaAuditAction: buildFormulaAuditAction(instance),
    updateErrorCheckingMenu: updateErrorCheckingMenu(instance),
    updateClearArrowsMenu: updateClearArrowsMenu(instance),
    applyWatchAction: buildWatchAction(instance),
    updateWatchMenu: updateWatchMenu(instance),
    applyCalcOptionAction: buildCalcOptionAction(instance),
    updateCalcOptionsMenu: updateCalcOptionsMenu(instance),
    applyDefinedNameAction: buildDefinedNameAction(instance),
    updateDefinedNamesMenu: updateDefinedNamesMenu(instance),
  };
}
