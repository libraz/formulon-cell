import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { type History, recordConditionalRulesChange } from '../commands/history.js';
import { type ConditionalRule, mutators, type SpreadsheetStore } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';

export interface CfRulesDialogDeps {
  host: HTMLElement;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so it always reads the live engine. */
  getWb: () => WorkbookHandle | null;
  /** Reads the active sheet at open time. The dialog scopes its rule list
   *  to a single sheet (the desktop-spreadsheet "Show formatting rules for: This Sheet"). */
  getActiveSheet: () => number;
  /** Called after a remove / clearAll mutation so the host can repaint the
   *  CF overlay. */
  onChanged?: () => void;
  /** Opens the companion "New Formatting Rule" dialog from Manage Rules. */
  onNewRule?: () => void;
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
      type: string;
      range: string;
    }
  | {
      source: 'session';
      index: number;
      priority: string;
      type: string;
      range: string;
    };

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

  const newRuleBtn = document.createElement('button');
  newRuleBtn.type = 'button';
  newRuleBtn.className = 'fc-cfrulesdlg__new';
  newRuleBtn.textContent = t.newRule;
  footer.appendChild(newRuleBtn);

  const clearAllBtn = document.createElement('button');
  clearAllBtn.type = 'button';
  clearAllBtn.className = 'fc-cfrulesdlg__clearall';
  clearAllBtn.textContent = t.clearAll;
  footer.appendChild(clearAllBtn);

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-cfrulesdlg__close';
  closeBtn.textContent = t.close;
  footer.appendChild(closeBtn);
  let selectedRuleIndex = 0;

  const focusRuleRow = (idx: number): void => {
    const rows = Array.from(tableWrap.querySelectorAll<HTMLTableRowElement>('tbody tr'));
    if (rows.length === 0) return;
    selectedRuleIndex = (idx + rows.length) % rows.length;
    for (const [rowIdx, row] of rows.entries()) {
      const selected = rowIdx === selectedRuleIndex;
      row.tabIndex = selected ? 0 : -1;
      row.setAttribute('aria-selected', selected ? 'true' : 'false');
      row.classList.toggle('fc-cfrulesdlg__row--selected', selected);
    }
    rows[selectedRuleIndex]?.focus({ preventScroll: true });
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
  shell.on(closeBtn, 'click', () => close());

  const renderTable = (): void => {
    tableWrap.replaceChildren();
    const wb = getWb();
    const sheet = getActiveSheet();
    const engineRules = wb?.getConditionalFormats(sheet) ?? [];
    const sessionRules = (deps.store?.getState().conditional.rules ?? [])
      .map((rule, index) => ({ rule, index }))
      .filter(({ rule }) => rule.range.sheet === sheet);
    const rules: ManagedRule[] = [
      ...engineRules.map<ManagedRule>((rule, index) => ({
        source: 'engine',
        index,
        priority: String(rule.priority),
        type: ruleTypeLabel(rule.type, t),
        range: formatSqref(rule.sqref),
      })),
      ...sessionRules.map<ManagedRule>(({ rule, index }, visibleIndex) => ({
        source: 'session',
        index,
        priority: String(engineRules.length + visibleIndex + 1),
        type: sessionRuleTypeLabel(rule, t),
        range: formatSqref([
          {
            firstRow: rule.range.r0,
            firstCol: rule.range.c0,
            lastRow: rule.range.r1,
            lastCol: rule.range.c1,
          },
        ]),
      })),
    ];
    if (rules.length === 0) {
      empty.hidden = false;
      clearAllBtn.disabled = true;
      return;
    }
    empty.hidden = true;
    clearAllBtn.disabled = false;
    const table = document.createElement('table');
    table.className = 'fc-cfrulesdlg__table';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    for (const label of [t.headerPriority, t.headerType, t.headerRange, t.headerActions]) {
      const th = document.createElement('th');
      th.textContent = label;
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    rules.forEach((managedRule, index) => {
      const row = document.createElement('tr');
      row.tabIndex = index === selectedRuleIndex ? 0 : -1;
      row.setAttribute('aria-selected', index === selectedRuleIndex ? 'true' : 'false');
      row.classList.toggle('fc-cfrulesdlg__row--selected', index === selectedRuleIndex);
      const prio = document.createElement('td');
      prio.textContent = managedRule.priority;
      const kind = document.createElement('td');
      kind.textContent = managedRule.type;
      const range = document.createElement('td');
      range.className = 'fc-cfrulesdlg__cell-range';
      range.textContent = managedRule.range;
      const actions = document.createElement('td');
      const moveUp = document.createElement('button');
      moveUp.type = 'button';
      moveUp.className = 'fc-cfrulesdlg__move-up';
      moveUp.textContent = t.moveUp;
      moveUp.disabled = managedRule.source !== 'session' || !deps.store || managedRule.index <= 0;
      moveUp.dataset.ruleIndex = String(index);
      moveUp.addEventListener('click', () => {
        if (managedRule.source !== 'session' || !deps.store || managedRule.index <= 0) return;
        recordConditionalRulesChange(deps.history ?? null, deps.store, () => {
          deps.store!.setState((state) => {
            const nextRules = [...state.conditional.rules];
            const [source] = nextRules.splice(managedRule.index, 1);
            if (!source) return state;
            nextRules.splice(managedRule.index - 1, 0, source);
            return { ...state, conditional: { rules: nextRules } };
          });
        });
        onChanged?.();
        selectedRuleIndex = Math.max(0, index - 1);
        renderTable();
        requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
      });
      const moveDown = document.createElement('button');
      moveDown.type = 'button';
      moveDown.className = 'fc-cfrulesdlg__move-down';
      moveDown.textContent = t.moveDown;
      moveDown.disabled =
        managedRule.source !== 'session' ||
        !deps.store ||
        managedRule.index >= deps.store.getState().conditional.rules.length - 1;
      moveDown.dataset.ruleIndex = String(index);
      moveDown.addEventListener('click', () => {
        if (
          managedRule.source !== 'session' ||
          !deps.store ||
          managedRule.index >= deps.store.getState().conditional.rules.length - 1
        ) {
          return;
        }
        recordConditionalRulesChange(deps.history ?? null, deps.store, () => {
          deps.store!.setState((state) => {
            const nextRules = [...state.conditional.rules];
            const [source] = nextRules.splice(managedRule.index, 1);
            if (!source) return state;
            nextRules.splice(managedRule.index + 1, 0, source);
            return { ...state, conditional: { rules: nextRules } };
          });
        });
        onChanged?.();
        selectedRuleIndex = Math.min(rules.length - 1, index + 1);
        renderTable();
        requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
      });
      const dup = document.createElement('button');
      dup.type = 'button';
      dup.className = 'fc-cfrulesdlg__duplicate';
      dup.textContent = t.duplicate;
      dup.disabled = managedRule.source !== 'session' || !deps.store;
      dup.dataset.ruleIndex = String(index);
      dup.addEventListener('click', () => {
        if (managedRule.source !== 'session' || !deps.store) return;
        recordConditionalRulesChange(deps.history ?? null, deps.store, () => {
          deps.store!.setState((state) => {
            const nextRules = [...state.conditional.rules];
            const source = nextRules[managedRule.index];
            if (!source) return state;
            nextRules.splice(managedRule.index + 1, 0, cloneSessionRule(source));
            return { ...state, conditional: { rules: nextRules } };
          });
        });
        onChanged?.();
        selectedRuleIndex = index + 1;
        renderTable();
        requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
      });
      const rm = document.createElement('button');
      rm.type = 'button';
      rm.className = 'fc-cfrulesdlg__remove';
      rm.textContent = t.remove;
      rm.dataset.ruleIndex = String(index);
      rm.addEventListener('click', () => {
        let changed = false;
        if (managedRule.source === 'engine') {
          if (!wb) return;
          changed = wb.removeConditionalFormatAt(sheet, managedRule.index);
        } else if (deps.store) {
          recordConditionalRulesChange(deps.history ?? null, deps.store, () => {
            mutators.removeConditionalRuleAt(deps.store!, managedRule.index);
          });
          changed = true;
        }
        if (changed) {
          onChanged?.();
          selectedRuleIndex = Math.min(index, Math.max(0, rules.length - 2));
          renderTable();
          requestAnimationFrame(() => focusRuleRow(selectedRuleIndex));
        }
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
      row.append(prio, kind, range, actions);
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
      recordConditionalRulesChange(deps.history ?? null, deps.store, () => {
        const before = deps.store!.getState().conditional.rules.length;
        mutators.clearConditionalRulesInRange(deps.store!, {
          sheet,
          r0: 0,
          c0: 0,
          r1: Number.MAX_SAFE_INTEGER,
          c1: Number.MAX_SAFE_INTEGER,
        });
        changed = changed || deps.store!.getState().conditional.rules.length !== before;
      });
    }
    if (changed) {
      onChanged?.();
      renderTable();
    }
  });

  const open = (): void => {
    resetArmed();
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
    newRuleBtn.textContent = t.newRule;
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
