import { type History, recordConditionalRulesChange } from '../commands/history.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { type ConditionalRule, mutators, type SpreadsheetStore } from '../store/store.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogButton, createDialogButton, createDialogShell } from './dialog-shell.js';

export interface CfRulesDialogDeps {
  host: HTMLElement;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so it always reads the live engine. */
  getWb: () => WorkbookHandle | null;
  /** Reads the active sheet at open time. The dialog scopes its rule list
   *  to a single sheet (the desktop-spreadsheet "Show formatting rules for: This Sheet"). */
  getActiveSheet: () => number;
  /** Reads the current grid selection for the "Current Selection" scope. */
  getSelectionRange?: () => Range | null;
  /** Called after a remove / clearAll mutation so the host can repaint the
   *  CF overlay. */
  onChanged?: () => void;
  /** Opens the companion "New Formatting Rule" dialog from Manage Rules. */
  onNewRule?: () => void;
  /** Opens the companion "Edit Formatting Rule" dialog for session rules. */
  onEditRule?: (ruleIndex: number) => void;
  /** Optional session-rule source. Rules created by the JS ribbon/Quick
   *  Analysis live in the store rather than the engine, but users expect the
   *  Manage Rules dialog to show the full active-sheet inventory. */
  store?: SpreadsheetStore;
  history?: History | null;
  strings?: Strings;
}

export interface CfRulesDialogHandle {
  open(): void;
  close(): void;
  refresh(): void;
  detach(): void;
}

function createCfRulesDialogButton(className: string, label: string): HTMLButtonElement {
  const button = createDialogButton({ label, baseClass: 'fc-cfrulesdlg__btn' });
  button.className = className;
  return button;
}

/** Maps the engine's CF rule type ordinal to a localized display label.
 *  Ordinals mirror `formulon::cf::RuleType`. */
const ruleTypeLabel = (type: number, t: Strings['cfRulesDialog']): string => {
  const labels: Readonly<Record<number, string>> = {
    0: t.ruleExpression,
    1: t.ruleCellIs,
    2: t.ruleColorScale,
    3: t.ruleDataBar,
    4: t.ruleIconSet,
    5: t.ruleTop10,
    6: t.ruleAboveAverage,
    7: t.ruleContainsText,
    8: t.ruleNotContainsText,
    9: t.ruleBeginsWith,
    10: t.ruleEndsWith,
    11: t.ruleContainsBlanks,
    12: t.ruleNotContainsBlanks,
    13: t.ruleContainsErrors,
    14: t.ruleNotContainsErrors,
    15: t.ruleTimePeriod,
    16: t.ruleDuplicateValues,
    17: t.ruleUniqueValues,
  };
  return labels[type] ?? `type=${type}`;
};

const colLetter = (col: number): string => {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
};

const formatSqref = (
  sqref: ReadonlyArray<{
    firstRow: number;
    firstCol: number;
    lastRow: number;
    lastCol: number;
  }>,
): string => {
  if (sqref.length === 0) return '';
  return sqref
    .map((r) => {
      const a = `${colLetter(r.firstCol)}${r.firstRow + 1}`;
      const b = `${colLetter(r.lastCol)}${r.lastRow + 1}`;
      return a === b ? a : `${a}:${b}`;
    })
    .join(' ');
};

const sessionRuleTypeLabel = (rule: ConditionalRule, t: Strings['cfRulesDialog']): string => {
  switch (rule.kind) {
    case 'formula':
      return t.ruleExpression;
    case 'cell-value':
      return t.ruleCellIs;
    case 'color-scale':
      return t.ruleColorScale;
    case 'data-bar':
      return t.ruleDataBar;
    case 'icon-set':
      return t.ruleIconSet;
    case 'top-bottom':
      return t.ruleTop10;
    case 'average':
      return t.ruleAboveAverage;
    case 'text-contains':
      return t.ruleContainsText;
    case 'date-occurring':
      return t.ruleTimePeriod;
    case 'duplicates':
      return t.ruleDuplicateValues;
    case 'unique':
      return t.ruleUniqueValues;
    case 'blanks':
      return t.ruleContainsBlanks;
    case 'non-blanks':
      return t.ruleNotContainsBlanks;
    case 'errors':
      return t.ruleContainsErrors;
    case 'no-errors':
      return t.ruleNotContainsErrors;
  }
};

const cloneSessionRule = (rule: ConditionalRule): ConditionalRule => {
  const out = { ...rule, range: { ...rule.range } } as ConditionalRule;
  if ('apply' in out) out.apply = { ...out.apply };
  if ('stops' in out) {
    out.stops = [...out.stops] as [string, string] | [string, string, string];
  }
  if ('thresholds' in out && out.thresholds) {
    out.thresholds = out.thresholds.map((point) => ({ ...point })) as typeof out.thresholds;
  }
  return out;
};

type ManagedRule =
  | {
      source: 'engine';
      index: number;
      priority: string;
      rule: string;
      format: string;
      range: string;
      stopIfTrue: boolean;
    }
  | {
      source: 'session';
      index: number;
      priority: string;
      rule: string;
      format: string;
      range: string;
      stopIfTrue: boolean;
    };

const rangesIntersect = (
  a: Range,
  b: { firstRow: number; firstCol: number; lastRow: number; lastCol: number },
): boolean => a.r0 <= b.lastRow && a.r1 >= b.firstRow && a.c0 <= b.lastCol && a.c1 >= b.firstCol;

/**
 * Engine-driven CF rule manager. Lists every rule on the active sheet via
 * `wb.getConditionalFormats(sheet)` and lets the user remove them
 * individually or clear the entire sheet. Authoring is deliberately
 * out of scope here — `addConditionalFormat` rejects visual rule types
 * upstream, and the existing JS-side `conditional-dialog.ts` already
 * covers basic rule creation. This is the "Manage Rules" surface.
 */
export function attachCfRulesDialog(deps: CfRulesDialogDeps): CfRulesDialogHandle {
  const { host, getWb, getActiveSheet, onChanged } = deps;
  let strings = deps.strings ?? defaultStrings;
  let t = strings.cfRulesDialog;

  const shell = createDialogShell({
    host,
    className: 'fc-cfrulesdlg',
    ariaLabel: t.title,
    onDismiss: () => close(),
  });
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-cfrulesdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-cfrulesdlg__body';
  panel.appendChild(body);

  const note = document.createElement('p');
  note.className = 'fc-cfrulesdlg__note';
  note.textContent = t.note;
  body.appendChild(note);

  const scopeRow = document.createElement('label');
  scopeRow.className = 'fc-cfrulesdlg__scope';
  const scopeText = document.createElement('span');
  scopeText.textContent = t.scopeLabel;
  const scopeSelect = createDialogSelect(
    [
      { value: 'selection', label: t.scopeCurrentSelection },
      { value: 'worksheet', label: t.scopeThisWorksheet },
    ],
    'worksheet',
    { className: 'fc-cfrulesdlg__scope-select' },
  );
  const setScopeOptionLabel = (value: string, label: string): void => {
    const option = Array.from(scopeSelect.options).find((item) => item.value === value);
    if (option) option.textContent = label;
  };
  scopeRow.append(scopeText, scopeSelect);
  body.appendChild(scopeRow);

  const commandBar = document.createElement('div');
  commandBar.className = 'fc-cfrulesdlg__commandbar';
  body.appendChild(commandBar);

  const makeCommandButton = (className: string, label: string): HTMLButtonElement => {
    return createCfRulesDialogButton(`fc-cfrulesdlg__cmd ${className}`, label);
  };

  const newRuleBtn = makeCommandButton('fc-cfrulesdlg__new', t.newRule);
  const editRuleBtn = makeCommandButton('fc-cfrulesdlg__edit', t.editRule);
  const duplicateBtn = makeCommandButton('fc-cfrulesdlg__duplicate', t.duplicate);
  const deleteBtn = makeCommandButton('fc-cfrulesdlg__delete', t.deleteRule);
  const moveUpBtn = makeCommandButton('fc-cfrulesdlg__move-up', t.moveUp);
  const moveDownBtn = makeCommandButton('fc-cfrulesdlg__move-down', t.moveDown);
  commandBar.append(newRuleBtn, editRuleBtn, duplicateBtn, deleteBtn, moveUpBtn, moveDownBtn);

  const tableWrap = document.createElement('div');
  tableWrap.className = 'fc-cfrulesdlg__tablewrap';
  body.appendChild(tableWrap);

  const empty = document.createElement('div');
  empty.className = 'fc-cfrulesdlg__empty';
  empty.textContent = t.empty;
  empty.hidden = true;
  body.appendChild(empty);

  const footer = document.createElement('div');
  footer.className = 'fc-cfrulesdlg__footer';
  panel.appendChild(footer);

  const clearAllBtn = appendDialogButton(footer, {
    label: t.clearAll,
    baseClass: 'fc-cfrulesdlg__clearall',
  });

  const closeBtn = appendDialogButton(footer, {
    label: t.close,
    baseClass: 'fc-cfrulesdlg__close',
  });
  let selectedRuleIndex = 0;
  let currentRules: ManagedRule[] = [];

  const setDisabledState = (
    button: HTMLElement & { disabled?: boolean },
    disabled: boolean,
    reason: string | null,
  ): void => {
    projectDisabledState(button, disabled, reason, { datasetKey: 'disabledReason' });
  };

  const updateCommandState = (): void => {
    const selected = currentRules[selectedRuleIndex] ?? null;
    const canEdit = selected?.source === 'session' && !!deps.store && !!deps.onEditRule;
    const canMutateSession = selected?.source === 'session' && !!deps.store;
    const engineReadOnlyReason = selected?.source === 'engine' ? t.editUnavailable : null;
    const noSelectionReason = selected ? null : t.selectRuleActionReason;
    const mutationUnavailableReason = selected && !deps.store ? t.note : null;
    setDisabledState(
      editRuleBtn,
      !canEdit,
      noSelectionReason ?? engineReadOnlyReason ?? mutationUnavailableReason,
    );
    setDisabledState(
      duplicateBtn,
      !canMutateSession,
      noSelectionReason ?? engineReadOnlyReason ?? mutationUnavailableReason,
    );
    setDisabledState(deleteBtn, !selected, t.selectRuleActionReason);
    const moveUpDisabled = !canMutateSession || selected.index <= 0;
    setDisabledState(
      moveUpBtn,
      moveUpDisabled,
      !selected
        ? t.selectRuleActionReason
        : (engineReadOnlyReason ??
            mutationUnavailableReason ??
            (selected.index <= 0 ? t.moveUpUnavailable : null)),
    );
    const moveDownDisabled =
      !canMutateSession ||
      selected.index >= (deps.store?.getState().conditional.rules.length ?? 0) - 1;
    setDisabledState(
      moveDownBtn,
      moveDownDisabled,
      !selected
        ? t.selectRuleActionReason
        : (engineReadOnlyReason ??
            mutationUnavailableReason ??
            (selected.index >= (deps.store?.getState().conditional.rules.length ?? 0) - 1
              ? t.moveDownUnavailable
              : null)),
    );
  };

  const removeManagedRule = (managedRule: ManagedRule, visualIndex: number): void => {
    const wb = getWb();
    const sheet = getActiveSheet();
    let changed = false;
    if (managedRule.source === 'engine') {
      if (!wb) return;
      changed = wb.removeConditionalFormatAt(sheet, managedRule.index);
    } else if (deps.store) {
      const store = deps.store;
      recordConditionalRulesChange(deps.history ?? null, store, () => {
        mutators.removeConditionalRuleAt(store, managedRule.index);
      });
      changed = true;
    }
    if (changed) {
      onChanged?.();
      selectedRuleIndex = Math.min(visualIndex, Math.max(0, currentRules.length - 2));
      renderTable();
      requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
    }
  };

  const duplicateManagedRule = (managedRule: ManagedRule, visualIndex: number): void => {
    if (managedRule.source !== 'session' || !deps.store) return;
    const store = deps.store;
    recordConditionalRulesChange(deps.history ?? null, store, () => {
      store.setState((state) => {
        const nextRules = [...state.conditional.rules];
        const source = nextRules[managedRule.index];
        if (!source) return state;
        nextRules.splice(managedRule.index + 1, 0, cloneSessionRule(source));
        return { ...state, conditional: { rules: nextRules } };
      });
    });
    onChanged?.();
    selectedRuleIndex = visualIndex + 1;
    renderTable();
    requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
  };

  const moveManagedRule = (managedRule: ManagedRule, visualIndex: number, delta: -1 | 1): void => {
    if (managedRule.source !== 'session' || !deps.store) return;
    const store = deps.store;
    const nextIndex = managedRule.index + delta;
    if (nextIndex < 0 || nextIndex >= store.getState().conditional.rules.length) return;
    recordConditionalRulesChange(deps.history ?? null, store, () => {
      store.setState((state) => {
        const nextRules = [...state.conditional.rules];
        const [source] = nextRules.splice(managedRule.index, 1);
        if (!source) return state;
        nextRules.splice(nextIndex, 0, source);
        return { ...state, conditional: { rules: nextRules } };
      });
    });
    onChanged?.();
    selectedRuleIndex = Math.max(0, Math.min(currentRules.length - 1, visualIndex + delta));
    renderTable();
    requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
  };

  const setManagedRuleStopIfTrue = (
    managedRule: ManagedRule,
    visualIndex: number,
    stopIfTrue: boolean,
  ): void => {
    if (managedRule.source !== 'session' || !deps.store) return;
    const store = deps.store;
    recordConditionalRulesChange(deps.history ?? null, store, () => {
      store.setState((state) => {
        const source = state.conditional.rules[managedRule.index];
        if (!source) return state;
        const nextRules = [...state.conditional.rules];
        nextRules[managedRule.index] = { ...source, stopIfTrue };
        return { ...state, conditional: { rules: nextRules } };
      });
    });
    onChanged?.();
    selectedRuleIndex = visualIndex;
    renderTable();
    requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
  };

  const focusRuleRow = (idx: number): void => {
    const rows = Array.from(tableWrap.querySelectorAll<HTMLTableRowElement>('tbody tr'));
    if (rows.length === 0) {
      updateCommandState();
      return;
    }
    selectedRuleIndex = (idx + rows.length) % rows.length;
    for (const [rowIdx, row] of rows.entries()) {
      const selected = rowIdx === selectedRuleIndex;
      row.tabIndex = selected ? 0 : -1;
      row.setAttribute('aria-selected', selected ? 'true' : 'false');
      row.classList.toggle('fc-cfrulesdlg__row--selected', selected);
    }
    rows[selectedRuleIndex]?.focus({ preventScroll: true });
    updateCommandState();
  };

  const onKey = (e: KeyboardEvent): void => {
    if (overlay.hidden) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };
  shell.on(document, 'keydown', onKey as EventListener);
  shell.on(newRuleBtn, 'click', () => {
    close();
    deps.onNewRule?.();
  });
  shell.on(editRuleBtn, 'click', () => {
    const selected = currentRules[selectedRuleIndex];
    if (selected?.source !== 'session') return;
    close();
    deps.onEditRule?.(selected.index);
  });
  shell.on(duplicateBtn, 'click', () => {
    const selected = currentRules[selectedRuleIndex];
    if (selected) duplicateManagedRule(selected, selectedRuleIndex);
  });
  shell.on(deleteBtn, 'click', () => {
    const selected = currentRules[selectedRuleIndex];
    if (selected) removeManagedRule(selected, selectedRuleIndex);
  });
  shell.on(moveUpBtn, 'click', () => {
    const selected = currentRules[selectedRuleIndex];
    if (selected) moveManagedRule(selected, selectedRuleIndex, -1);
  });
  shell.on(moveDownBtn, 'click', () => {
    const selected = currentRules[selectedRuleIndex];
    if (selected) moveManagedRule(selected, selectedRuleIndex, 1);
  });
  shell.on(closeBtn, 'click', () => close());
  shell.on(scopeSelect, 'change', () => {
    selectedRuleIndex = 0;
    renderTable();
  });

  const renderTable = (): void => {
    tableWrap.replaceChildren();
    const wb = getWb();
    const sheet = getActiveSheet();
    const scopeSelection = scopeSelect.value === 'selection' ? deps.getSelectionRange?.() : null;
    const engineRules = (wb?.getConditionalFormats(sheet) ?? [])
      .map((rule, index) => ({ rule, index }))
      .filter(
        ({ rule }) =>
          !scopeSelection || rule.sqref.some((range) => rangesIntersect(scopeSelection, range)),
      );
    const sessionRules = (deps.store?.getState().conditional.rules ?? [])
      .map((rule, index) => ({ rule, index }))
      .filter(
        ({ rule }) =>
          rule.range.sheet === sheet &&
          !rule.engineId &&
          (!scopeSelection ||
            rangesIntersect(scopeSelection, {
              firstRow: rule.range.r0,
              firstCol: rule.range.c0,
              lastRow: rule.range.r1,
              lastCol: rule.range.c1,
            })),
      );
    const rules: ManagedRule[] = [
      ...engineRules.map<ManagedRule>(({ rule, index }) => ({
        source: 'engine',
        index,
        priority: String(rule.priority),
        rule: ruleTypeLabel(rule.type, t),
        format: ruleTypeLabel(rule.type, t),
        range: formatSqref(rule.sqref),
        stopIfTrue: rule.stopIfTrue,
      })),
      ...sessionRules.map<ManagedRule>(({ rule, index }, visibleIndex) => ({
        source: 'session',
        index,
        priority: String(engineRules.length + visibleIndex + 1),
        rule: sessionRuleTypeLabel(rule, t),
        format: sessionRuleTypeLabel(rule, t),
        range: formatSqref([
          {
            firstRow: rule.range.r0,
            firstCol: rule.range.c0,
            lastRow: rule.range.r1,
            lastCol: rule.range.c1,
          },
        ]),
        stopIfTrue: rule.stopIfTrue === true,
      })),
    ];
    currentRules = rules;
    if (rules.length === 0) {
      empty.hidden = false;
      setDisabledState(clearAllBtn, true, t.clearAllRequiresRules);
      updateCommandState();
      return;
    }
    empty.hidden = true;
    setDisabledState(clearAllBtn, false, null);
    const table = document.createElement('table');
    table.className = 'fc-cfrulesdlg__table';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    for (const label of [
      t.headerRule,
      t.headerFormat,
      t.headerAppliesTo,
      t.headerStopIfTrue,
      t.headerActions,
    ]) {
      const th = document.createElement('th');
      th.textContent = label;
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    rules.forEach((managedRule, index) => {
      const row = document.createElement('tr');
      row.dataset.ruleSource = managedRule.source;
      row.tabIndex = index === selectedRuleIndex ? 0 : -1;
      row.setAttribute('aria-selected', index === selectedRuleIndex ? 'true' : 'false');
      row.classList.toggle('fc-cfrulesdlg__row--selected', index === selectedRuleIndex);
      row.classList.toggle('fc-cfrulesdlg__row--readonly', managedRule.source === 'engine');
      row.addEventListener('click', () => focusRuleRow(index));
      const kind = document.createElement('td');
      kind.className = 'fc-cfrulesdlg__cell-rule';
      kind.textContent = managedRule.rule;
      const format = document.createElement('td');
      format.className = 'fc-cfrulesdlg__cell-format';
      format.textContent = managedRule.format;
      if (managedRule.source === 'engine') {
        const badge = document.createElement('span');
        badge.className = 'fc-cfrulesdlg__readonly-badge';
        badge.textContent = t.readOnlyRule;
        badge.title = t.editUnavailable;
        badge.setAttribute('aria-label', t.editUnavailable);
        format.append(' ', badge);
      }
      const range = document.createElement('td');
      range.className = 'fc-cfrulesdlg__cell-range';
      range.textContent = managedRule.range;
      const stop = document.createElement('td');
      stop.className = 'fc-cfrulesdlg__cell-stop';
      const stopCheckbox = document.createElement('input');
      stopCheckbox.type = 'checkbox';
      stopCheckbox.checked = managedRule.stopIfTrue;
      setDisabledState(
        stopCheckbox,
        managedRule.source !== 'session' || !deps.store,
        managedRule.source === 'engine' ? t.stopIfTrueUnavailable : !deps.store ? t.note : null,
      );
      stopCheckbox.addEventListener('change', () => {
        setManagedRuleStopIfTrue(managedRule, index, stopCheckbox.checked);
      });
      stop.setAttribute('aria-label', t.headerStopIfTrue);
      stop.appendChild(stopCheckbox);
      const actions = document.createElement('td');
      actions.className = 'fc-cfrulesdlg__cell-actions';
      const moveUp = createCfRulesDialogButton('fc-cfrulesdlg__row-move-up', t.moveUp);
      setDisabledState(
        moveUp,
        managedRule.source !== 'session' || !deps.store || managedRule.index <= 0,
        managedRule.source === 'engine'
          ? t.editUnavailable
          : !deps.store
            ? t.note
            : managedRule.index <= 0
              ? t.moveUpUnavailable
              : null,
      );
      moveUp.dataset.ruleIndex = String(index);
      moveUp.addEventListener('click', () => {
        moveManagedRule(managedRule, index, -1);
      });
      const moveDown = createCfRulesDialogButton('fc-cfrulesdlg__row-move-down', t.moveDown);
      setDisabledState(
        moveDown,
        managedRule.source !== 'session' ||
          !deps.store ||
          managedRule.index >= deps.store.getState().conditional.rules.length - 1,
        managedRule.source === 'engine'
          ? t.editUnavailable
          : !deps.store
            ? t.note
            : managedRule.index >= deps.store.getState().conditional.rules.length - 1
              ? t.moveDownUnavailable
              : null,
      );
      moveDown.dataset.ruleIndex = String(index);
      moveDown.addEventListener('click', () => {
        moveManagedRule(managedRule, index, 1);
      });
      const dup = createCfRulesDialogButton('fc-cfrulesdlg__row-duplicate', t.duplicate);
      setDisabledState(
        dup,
        managedRule.source !== 'session' || !deps.store,
        managedRule.source === 'engine' ? t.editUnavailable : !deps.store ? t.note : null,
      );
      dup.dataset.ruleIndex = String(index);
      dup.addEventListener('click', () => {
        duplicateManagedRule(managedRule, index);
      });
      const rm = createCfRulesDialogButton('fc-cfrulesdlg__remove', t.remove);
      rm.dataset.ruleIndex = String(index);
      rm.addEventListener('click', () => {
        removeManagedRule(managedRule, index);
      });
      row.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          focusRuleRow(index + 1);
        } else if (e.key === 'ArrowUp') {
          e.preventDefault();
          focusRuleRow(index - 1);
        } else if (e.key === 'Home') {
          e.preventDefault();
          focusRuleRow(0);
        } else if (e.key === 'End') {
          e.preventDefault();
          focusRuleRow(rules.length - 1);
        } else if (e.key === 'Delete' || e.key === 'Backspace') {
          e.preventDefault();
          rm.click();
        } else if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          rm.focus();
        }
      });
      actions.append(moveUp, moveDown, dup, rm);
      row.append(kind, format, range, stop, actions);
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
    tableWrap.appendChild(table);
    selectedRuleIndex = Math.min(selectedRuleIndex, rules.length - 1);
    focusRuleRow(selectedRuleIndex);
  };

  // Two-step "Clear all": first click arms the button, second click within
  // 3 seconds confirms the destruction. Avoids a separate confirm() dialog.
  let armed = false;
  let armTimer: ReturnType<typeof setTimeout> | null = null;
  const resetArmed = (): void => {
    if (armTimer) {
      clearTimeout(armTimer);
      armTimer = null;
    }
    armed = false;
    clearAllBtn.classList.remove('fc-cfrulesdlg__clearall--armed');
    clearAllBtn.textContent = t.clearAll;
  };
  shell.on(clearAllBtn, 'click', () => {
    const wb = getWb();
    const sheet = getActiveSheet();
    if (!wb && !deps.store) return;
    if (!armed) {
      armed = true;
      clearAllBtn.classList.add('fc-cfrulesdlg__clearall--armed');
      clearAllBtn.textContent = t.clearAllConfirm;
      armTimer = setTimeout(resetArmed, 3000);
      return;
    }
    resetArmed();
    let changed = wb?.clearConditionalFormats(sheet) ?? false;
    if (deps.store) {
      const store = deps.store;
      recordConditionalRulesChange(deps.history ?? null, store, () => {
        const before = store.getState().conditional.rules.length;
        mutators.clearConditionalRulesInRange(store, {
          sheet,
          r0: 0,
          c0: 0,
          r1: Number.MAX_SAFE_INTEGER,
          c1: Number.MAX_SAFE_INTEGER,
        });
        changed = changed || store.getState().conditional.rules.length !== before;
      });
    }
    if (changed) {
      onChanged?.();
      renderTable();
    }
  });

  const open = (): void => {
    resetArmed();
    scopeSelect.value = 'worksheet';
    renderTable();
    shell.open();
  };

  const close = (): void => {
    resetArmed();
    shell.close();
  };

  const refresh = (): void => {
    strings = deps.strings ?? defaultStrings;
    t = strings.cfRulesDialog;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    note.textContent = t.note;
    empty.textContent = t.empty;
    scopeText.textContent = t.scopeLabel;
    setScopeOptionLabel('selection', t.scopeCurrentSelection);
    setScopeOptionLabel('worksheet', t.scopeThisWorksheet);
    newRuleBtn.textContent = t.newRule;
    editRuleBtn.textContent = t.editRule;
    duplicateBtn.textContent = t.duplicate;
    deleteBtn.textContent = t.deleteRule;
    moveUpBtn.textContent = t.moveUp;
    moveDownBtn.textContent = t.moveDown;
    closeBtn.textContent = t.close;
    if (!armed) clearAllBtn.textContent = t.clearAll;
    if (!overlay.hidden) renderTable();
  };

  return {
    open,
    close,
    refresh,
    detach() {
      resetArmed();
      shell.dispose();
    },
  };
}
