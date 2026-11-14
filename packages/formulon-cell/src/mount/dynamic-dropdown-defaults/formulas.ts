import type { AutoSumFunction } from '../../commands/auto-sum.js';
import { canExecuteBuiltIn } from '../../commands/built-in-command-policy.js';
import { circularReferenceEntries } from '../../commands/circular-references.js';
import {
  COMPATIBILITY_VERSIONS,
  chooseCompatibilityVersion,
  compatibilityVersionLabel,
  compatibilityVersionRow,
} from '../../commands/compatibility-version.js';
import { DATA_TABLE_COMMAND } from '../../commands/data-table.js';
import { cellValueIsFormulaError } from '../../commands/error-indicators.js';
import { namedRangeWrite } from '../../commands/named-range-actions.js';
import {
  definedNameExists,
  insertDefinedNameFormula,
  listDefinedNames,
  type PlannedDefinedName,
  planDefinedNamesFromSelection,
  upsertDefinedName,
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
import { addrKey, formatA1Cell, quoteSheetName } from '../../engine/address.js';
import {
  type CompatibilityVersionCode,
  calcModeLeavesStaleValues,
} from '../../engine/calc-mode.js';
import type { Range } from '../../engine/types.js';
import { isMacHost } from '../../extensions/ui-options.js';
import { openDataTableDialogFrom } from '../../interact/data-table-dialog.js';
import { openDefineNameDialogFrom, openNamedRangeDialogFrom } from '../../interact/dialog-entry.js';
import { followHyperlink } from '../../interact/hyperlink-follow.js';
import { openMacGoalSeekDialog } from '../../interact/mac-goal-seek-dialog.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { mutators } from '../../store/store.js';
import { showCommandRefusal } from '../../toolbar/dialogs/command-refusal.js';
import { showCreateNamesDialog } from '../../toolbar/dialogs/create-names.js';
import { showDefinedNamePickerDialog } from '../../toolbar/dialogs/defined-name-picker.js';
import { showConfirm } from '../../toolbar/dialogs/prompt.js';
import { createRibbonButton } from '../../toolbar/ribbon/button.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import {
  menuIconSpacer,
  menuPresetButton,
  menuSeparator,
} from '../../toolbar/ribbon/menus/general.js';
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
      return;
    }
    const version = /^compat-([123])$/.exec(action)?.[1];
    if (version) {
      chooseCompatibilityVersion(
        instance.workbook,
        instance.history,
        Number(version) as CompatibilityVersionCode,
      );
      instance.host.focus();
      return;
    }
    if (action === 'format-stale') {
      if (!instance.workbook.staleValuesSupported()) return;
      mutators.setFormatStaleValues(
        instance.store,
        !instance.store.getState().ui.formatStaleValues,
      );
      instance.host.focus();
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
    const mode = instance.workbook.calcMode();
    const current = calcOptionForMode(mode);
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]')) {
      if (button.dataset.calcOption?.startsWith('compat-')) continue;
      const active = button.dataset.calcOption === current;
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('fc-tb__menu-item--active', active);
    }
    const unavailable = instance.i18n.strings.backstage.commandUnavailable;
    const stale = menu.querySelector<HTMLButtonElement>('[data-calc-option="format-stale"]');
    if (stale) {
      // Under Automatic it stays off with no reason, as on the desktop.
      const supported = instance.workbook.staleValuesSupported();
      setMenuControlDisabled(
        stale,
        !supported || !calcModeLeavesStaleValues(mode),
        supported ? undefined : unavailable,
      );
      stale.setAttribute('aria-checked', String(instance.store.getState().ui.formatStaleValues));
    }
    const version = instance.workbook.compatibilityVersion();
    const trigger = menu.querySelector<HTMLButtonElement>('[data-format-submenu="compat-version"]');
    if (trigger) setMenuControlDisabled(trigger, version === null, unavailable);
    for (const choice of COMPATIBILITY_VERSIONS) {
      const button = menu.querySelector<HTMLButtonElement>(`[data-calc-option="compat-${choice}"]`);
      if (!button) continue;
      const row = compatibilityVersionRow(choice, version);
      button.hidden = !row.listed;
      button.setAttribute('aria-checked', String(row.checked));
      const text = button.querySelector('.fc-tb__menu-item__text');
      if (text)
        text.textContent = compatibilityVersionLabel(
          instance.i18n.strings.ribbonMenu,
          choice,
          row.shown,
        );
      setMenuControlDisabled(button, version === null || row.locked, unavailable);
    }
    // The host has no help page to open.
    const learnMore = menu.querySelector<HTMLButtonElement>(
      '[data-calc-option="compat-learn-more"]',
    );
    if (learnMore) setMenuControlDisabled(learnMore, true, unavailable);
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
    projectCircularReferences(instance, menu);
  };

/** Both profiles' Circular References row and panel; the Mac keys them by command id. */
const CIRCULAR_SUBMENU = [
  '[data-format-submenu="circular"]',
  '[data-format-submenu="mac.formulas.errorCheck.circular"]',
].join(', ');
const CIRCULAR_PANEL = [
  '[data-format-panel="circular"]',
  '[data-format-panel="mac.formulas.errorCheck.circular"]',
].join(', ');

/** Lists the cycle members as the menu opens; the row is off when there are none. */
const projectCircularReferences = (instance: SpreadsheetInstance, menu: HTMLElement): void => {
  const trigger = menu.querySelector<HTMLButtonElement>(CIRCULAR_SUBMENU);
  const panel = menu.querySelector<HTMLElement>(CIRCULAR_PANEL);
  if (!trigger || !panel) return;
  const supported = instance.workbook.circularReferencesSupported();
  const state = instance.store.getState();
  const entries = supported
    ? circularReferenceEntries(instance.workbook, state.data.sheetIndex)
    : [];
  const active = state.selection.active;
  const ticked =
    entries.find(
      (entry) =>
        entry.sheet === active.sheet && entry.row === active.row && entry.col === active.col,
    ) ?? entries[0];
  const mac = menu.classList.contains('fc-tb__menu--mac');
  panel.replaceChildren(
    ...entries.map((entry) => {
      const button = menuPresetButton(
        entry.label,
        'circularReference',
        `${entry.sheet}:${entry.row}:${entry.col}`,
        menuIconSpacer(),
      );
      if (mac) button.classList.add('fc-tb__menu-item--mac');
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(entry === ticked));
      button.tabIndex = -1;
      return button;
    }),
  );
  setMenuControlDisabled(
    trigger,
    entries.length === 0,
    supported ? undefined : instance.i18n.strings.backstage.commandUnavailable,
  );
};

const buildSelectCircularReference =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['selectCircularReference'] =>
  (cell) => {
    const [sheet, row, col] = cell.split(':').map(Number);
    if (sheet === undefined || row === undefined || col === undefined) return;
    const target = `#${quoteSheetName(instance.workbook.sheetName(sheet))}!${formatA1Cell(row, col, true)}`;
    followHyperlink({ store: instance.store, wb: instance.workbook }, target);
    instance.host.focus();
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

/** Create from Selection: the Create Names dialog, then the names from each checked edge as one
 *  step. */
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
    // The corner cell the top row and left column share decides neither.
    initial: {
      top: holdsText(instance, {
        ...range,
        r1: range.r0,
        c0: range.c1 > range.c0 ? range.c0 + 1 : range.c0,
      }),
      left: holdsText(instance, {
        ...range,
        c1: range.c0,
        r0: range.r1 > range.r0 ? range.r0 + 1 : range.r0,
      }),
      bottom: false,
      right: false,
    },
  });
  if (!edges) return;
  const planned = planDefinedNamesFromSelection(
    instance.store.getState(),
    instance.workbook,
    range,
    edges,
  );
  if (!planned.ok) {
    await showCommandRefusal({
      host: instance.host,
      strings,
      title: t.createNamesTitle,
      reason: 'too-large',
    });
    return;
  }
  // A name the workbook already holds is replaced only on Yes. Provisional: asked before any
  // write rather than as each name is reached.
  const names: PlannedDefinedName[] = [];
  for (const entry of planned.names) {
    if (
      definedNameExists(instance.workbook, entry.name) &&
      !(await showConfirm({
        title: t.createNamesTitle,
        message: t.createNamesReplace.replace('{name}', entry.name),
        okLabel: t.createNamesReplaceYes,
        cancelLabel: t.createNamesReplaceNo,
        alert: isMacHost(instance.host),
      }))
    ) {
      continue;
    }
    names.push(entry);
  }
  if (names.length === 0) return;
  const written = namedRangeWrite(
    {
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      origin: 'ribbon',
      commandId: 'namedRanges',
    },
    () => {
      for (const { name, formula } of names) {
        const result = upsertDefinedName(instance.workbook, name, formula);
        if (!result.ok) return result.reason;
      }
      return null;
    },
  );
  const failure = written.ok ? written.value : null;
  if (failure !== null && failure !== 'unsupported') {
    await showCommandRefusal({
      host: instance.host,
      strings,
      title: t.createNamesTitle,
      reason: 'engine-failed',
    });
  }
}

/** Writes `=name` into the active cell (after any allowed-range unlock), then redraws and refocuses. */
const insertNameInActiveCell = (instance: SpreadsheetInstance, name: string): void => {
  const { active } = instance.store.getState().selection;
  const cell = {
    sheet: active.sheet,
    r0: active.row,
    c0: active.col,
    r1: active.row,
    c1: active.col,
  };
  runAfterAllowedRangeUnlock(instance.store, [cell], () => {
    insertDefinedNameFormula(instance.store.getState(), instance.workbook, name, instance.store);
    mutators.refreshActiveSheetCells(instance.store, instance.workbook);
    instance.host.focus();
  });
};

const INSERT_NAME_ACTION_PREFIX = 'insert:';

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
    if (action.startsWith(INSERT_NAME_ACTION_PREFIX)) {
      insertNameInActiveCell(instance, action.slice(INSERT_NAME_ACTION_PREFIX.length));
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
      insertNameInActiveCell(instance, selected);
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

/** Collation for the name list: case-insensitive, Latin before kana before kanji. */
const nameCollator = new Intl.Collator('ja', { sensitivity: 'base' });

/** Names the Use in Formula menu lists: workbook names plus the active sheet's own, once each. */
const listedNamesForActiveSheet = (instance: SpreadsheetInstance): string[] => {
  const sheet = instance.store.getState().selection.active.sheet;
  const seen = new Map<string, string>();
  for (const entry of listDefinedNames(instance.workbook)) {
    if (entry.hidden || (entry.localSheetId !== -1 && entry.localSheetId !== sheet)) continue;
    const key = entry.name.trim().toLowerCase();
    if (entry.localSheetId === sheet || !seen.has(key)) seen.set(key, entry.name);
  }
  return [...seen.values()].sort(nameCollator.compare);
};

/** Rebuilds the name rows and their separator ahead of Paste Names each time the menu opens. */
const updateUseInFormulaMenu =
  (instance: SpreadsheetInstance): NonNullable<DynamicDropdownsCtx['updateUseInFormulaMenu']> =>
  (menu) => {
    for (const stale of menu.querySelectorAll('[data-defined-name-insert]')) stale.remove();
    const paste = menu.querySelector('[data-ribbon-command="mac.formulas.useInFormula.paste"]');
    const names = listedNamesForActiveSheet(instance);
    if (!paste || names.length === 0) return;
    const rows = names.map((name) =>
      createRibbonButton({
        className: 'fc-tb__menu-item fc-tb__menu-item--mac',
        role: 'menuitem',
        title: name,
        ariaLabel: name,
        text: name,
        dataset: {
          definedNameAction: `${INSERT_NAME_ACTION_PREFIX}${name}`,
          definedNameInsert: name,
        },
      }),
    );
    const separator = menuSeparator();
    separator.dataset.definedNameInsert = '';
    paste.before(...rows, separator);
  };

type FormulasDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyAutoSumFormula'
  | 'applyFormulaAuditAction'
  | 'updateErrorCheckingMenu'
  | 'selectCircularReference'
  | 'updateClearArrowsMenu'
  | 'applyWatchAction'
  | 'updateWatchMenu'
  | 'applyCalcOptionAction'
  | 'updateCalcOptionsMenu'
  | 'applyWhatIfAction'
  | 'updateWhatIfMenu'
  | 'applyDefinedNameAction'
  | 'updateDefinedNamesMenu'
  | 'updateUseInFormulaMenu'
>;

export function createFormulasDropdownDefaults(
  instance: SpreadsheetInstance,
): FormulasDropdownDefaults {
  return {
    applyAutoSumFormula: buildAutoSumFormula(instance),
    applyFormulaAuditAction: buildFormulaAuditAction(instance),
    updateErrorCheckingMenu: updateErrorCheckingMenu(instance),
    selectCircularReference: buildSelectCircularReference(instance),
    updateClearArrowsMenu: updateClearArrowsMenu(instance),
    applyWatchAction: buildWatchAction(instance),
    updateWatchMenu: updateWatchMenu(instance),
    applyCalcOptionAction: buildCalcOptionAction(instance),
    updateCalcOptionsMenu: updateCalcOptionsMenu(instance),
    applyWhatIfAction: buildWhatIfAction(instance),
    updateWhatIfMenu: updateWhatIfMenu(instance),
    applyDefinedNameAction: buildDefinedNameAction(instance),
    updateDefinedNamesMenu: updateDefinedNamesMenu(instance),
    updateUseInFormulaMenu: updateUseInFormulaMenu(instance),
  };
}
