import type { AutoSumFunction } from '../../commands/auto-sum.js';
import { canExecuteBuiltIn } from '../../commands/built-in-command-policy.js';
import { DATA_TABLE_COMMAND } from '../../commands/data-table.js';
import { cellValueIsFormulaError } from '../../commands/error-indicators.js';
import { namedRangeWrite } from '../../commands/named-range-actions.js';
import {
  createDefinedNamesFromSelection,
  insertDefinedNameFormula,
  listDefinedNames,
} from '../../commands/named-ranges.js';
import { localizedDecisionReason } from '../../commands/permission-reasons.js';
import { runAfterAllowedRangeUnlock } from '../../commands/protection.js';
import { executeRibbonFormulaAuditingAction } from '../../commands/ribbon-formula-auditing.js';
import { clearTraceArrowsByKind } from '../../commands/traces.js';
import {
  clearWatchedCells,
  recordWatchesChange,
  unwatchCell,
  watchRange,
} from '../../commands/watch.js';
import { addrKey } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import { openDataTableDialogFrom } from '../../interact/data-table-dialog.js';
import { openDefineNameDialogFrom, openNamedRangeDialogFrom } from '../../interact/dialog-entry.js';
import { openMacGoalSeekDialog } from '../../interact/mac-goal-seek-dialog.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { mutators } from '../../store/store.js';
import { showCommandRefusal } from '../../toolbar/dialogs/command-refusal.js';
import { showCreateNamesDialog } from '../../toolbar/dialogs/create-names.js';
import { showDefinedNamePickerDialog } from '../../toolbar/dialogs/defined-name-picker.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { handleAutoSum } from '../../wrappers/toolbar-actions.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';

const buildAutoSumFormula =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyAutoSumFormula'] =>
  (fn) => {
    if (fn === 'MORE') {
      instance.openFunctionArguments();
      return;
    }
    // Focus first: an editor or unlock prompt opened by AutoSum takes it over.
    instance.host.focus();
    handleAutoSum(instance, fn as AutoSumFunction);
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

const buildWhatIfAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyWhatIfAction'] =>
  (action) => {
    if (action === 'goal-seek') openMacGoalSeekDialog(instance);
    else if (action === 'data-table') openDataTableDialogFrom(instance, 'ribbon');
  };

const updateWhatIfMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateWhatIfMenu'] =>
  (menu) => {
    const dataTable = menu.querySelector<HTMLButtonElement>('[data-what-if="data-table"]');
    if (!dataTable) return;
    if (!instance.workbook.capabilities.dataTables) {
      setMenuControlDisabled(
        dataTable,
        true,
        instance.i18n.strings.ribbonMenu.whatIfDataTableUnavailable,
      );
      return;
    }
    // A protected sheet, or a selection over a table, refuses a Data Table.
    const decision = canExecuteBuiltIn(instance.store, DATA_TABLE_COMMAND, 'ribbon');
    setMenuControlDisabled(
      dataTable,
      !decision.allowed,
      decision.allowed ? undefined : localizedDecisionReason(decision, instance.i18n.strings),
    );
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
      macNa:
        (instance.host.closest<HTMLElement>('.fc-host') ?? instance.host).dataset.fcPlatform ===
        'mac',
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

const CREATE_NAMES_SOURCES = [
  ['top', 'top-row'],
  ['left', 'left-column'],
  ['bottom', 'bottom-row'],
  ['right', 'right-column'],
] as const;

/** Whether any cell of `range` holds text, which the Create Names dialog reads as labels. */
const holdsText = (instance: SpreadsheetInstance, range: Range): boolean => {
  const cells = instance.store.getState().data.cells;
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      if (cells.get(addrKey({ sheet: range.sheet, row, col }))?.value.kind === 'text') return true;
    }
  }
  return false;
};

/** Create from Selection: the Create Names dialog, then names from each checked edge as one step.
 *  An edge's names leave out the other checked edges, so with Top row and Left column checked the
 *  column names start after the left column and the row names after the top row. */
async function createNamesFromSelection(instance: SpreadsheetInstance): Promise<void> {
  const strings = instance.i18n.strings;
  const t = strings.ribbonMenu;
  const range = instance.store.getState().selection.range;
  // Top row and Left column start checked when that edge holds text; the desktop's own
  // pre-fill rule is provisional until measured.
  const edges = await showCreateNamesDialog({
    host: instance.host,
    labels: {
      title: t.createNamesTitle,
      createIn: t.createNamesIn,
      top: t.createNamesTop,
      left: t.createNamesLeft,
      bottom: t.createNamesBottom,
      right: t.createNamesRight,
      ok: strings.pageSetup.ok,
      cancel: strings.pageSetup.cancel,
    },
    initial: {
      top: holdsText(instance, { ...range, r1: range.r0 }),
      left: holdsText(instance, { ...range, c1: range.c0 }),
      bottom: false,
      right: false,
    },
  });
  if (!edges) return;
  const narrowed = (source: (typeof CREATE_NAMES_SOURCES)[number][1]): Range => {
    const rows = source === 'top-row' || source === 'bottom-row';
    return {
      ...range,
      r0: !rows && edges.top ? range.r0 + 1 : range.r0,
      r1: !rows && edges.bottom ? range.r1 - 1 : range.r1,
      c0: rows && edges.left ? range.c0 + 1 : range.c0,
      c1: rows && edges.right ? range.c1 - 1 : range.c1,
    };
  };
  const written = namedRangeWrite(
    {
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      origin: 'ribbon',
      commandId: 'namedRanges',
    },
    () => {
      for (const [edge, source] of CREATE_NAMES_SOURCES) {
        if (!edges[edge]) continue;
        const state = instance.store.getState();
        const result = createDefinedNamesFromSelection(
          { ...state, selection: { ...state.selection, range: narrowed(source) } },
          instance.workbook,
          source,
        );
        if (!result.ok) return result;
      }
      return null;
    },
  );
  const reason = written.ok && written.value ? written.value.reason : null;
  if (reason === 'too-large' || reason === 'engine-failed') {
    await showCommandRefusal({ host: instance.host, strings, title: t.createNamesTitle, reason });
  }
}

const buildDefinedNameAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyDefinedNameAction'] =>
  async (action) => {
    if (action === 'define') {
      openDefineNameDialogFrom(instance.store, { origin: 'ribbon', commandId: 'defineName' });
      return;
    }
    if (action === 'manager') {
      openNamedRangeDialogFrom(instance.store, { origin: 'ribbon', commandId: 'namedRanges' });
      return;
    }
    if (action === 'create-from-selection') {
      await createNamesFromSelection(instance);
      instance.host.focus();
      return;
    }
    if (action === 'use-formula') {
      const names = listDefinedNames(instance.workbook).filter((entry) => !entry.hidden);
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
      const { active } = instance.store.getState().selection;
      const cell = {
        sheet: active.sheet,
        r0: active.row,
        c0: active.col,
        r1: active.row,
        c1: active.col,
      };
      runAfterAllowedRangeUnlock(instance.store, [cell], () => {
        insertDefinedNameFormula(
          instance.store.getState(),
          instance.workbook,
          selected,
          instance.store,
        );
        mutators.refreshActiveSheetCells(instance.store, instance.workbook);
        instance.host.focus();
      });
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
  | 'applyWhatIfAction'
  | 'updateWhatIfMenu'
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
    applyWhatIfAction: buildWhatIfAction(instance),
    updateWhatIfMenu: updateWhatIfMenu(instance),
    applyDefinedNameAction: buildDefinedNameAction(instance),
    updateDefinedNamesMenu: updateDefinedNamesMenu(instance),
  };
}
