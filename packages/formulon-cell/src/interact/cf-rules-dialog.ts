import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
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
  strings?: Strings;
}

export interface CfRulesDialogHandle {
  open(): void;
  close(): void;
  refresh(): void;
  detach(): void;
}

/** Maps the engine's CF rule type ordinal to a short display label.
 *  Ordinals mirror `formulon::cf::RuleType`; visual rules are surfaced
 *  but flagged as read-only in the action column. */
const RULE_TYPE_LABELS: Readonly<Record<number, string>> = {
  0: 'expression',
  1: 'cellIs',
  2: 'colorScale',
  3: 'dataBar',
  4: 'iconSet',
  5: 'top10',
  6: 'aboveAverage',
  7: 'containsText',
  8: 'notContainsText',
  9: 'beginsWith',
  10: 'endsWith',
  11: 'containsBlanks',
  12: 'notContainsBlanks',
  13: 'containsErrors',
  14: 'notContainsErrors',
  15: 'timePeriod',
  16: 'duplicateValues',
  17: 'uniqueValues',
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
  shell.on(closeBtn, 'click', () => close());

  const renderTable = (): void => {
    tableWrap.replaceChildren();
    const wb = getWb();
    const sheet = getActiveSheet();
    const rules = wb?.getConditionalFormats(sheet) ?? [];
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
    rules.forEach((rule, index) => {
      const row = document.createElement('tr');
      row.tabIndex = index === selectedRuleIndex ? 0 : -1;
      row.setAttribute('aria-selected', index === selectedRuleIndex ? 'true' : 'false');
      row.classList.toggle('fc-cfrulesdlg__row--selected', index === selectedRuleIndex);
      const prio = document.createElement('td');
      prio.textContent = String(rule.priority);
      const kind = document.createElement('td');
      kind.textContent = RULE_TYPE_LABELS[rule.type] ?? `type=${rule.type}`;
      const range = document.createElement('td');
      range.className = 'fc-cfrulesdlg__cell-range';
      range.textContent = formatSqref(rule.sqref);
      const actions = document.createElement('td');
      const rm = document.createElement('button');
      rm.type = 'button';
      rm.className = 'fc-cfrulesdlg__remove';
      rm.textContent = t.remove;
      rm.dataset.ruleIndex = String(index);
      rm.addEventListener('click', () => {
        if (!wb) return;
        if (wb.removeConditionalFormatAt(sheet, index)) {
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
      actions.appendChild(rm);
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
    if (!wb) return;
    if (!armed) {
      armed = true;
      clearAllBtn.classList.add('fc-cfrulesdlg__clearall--armed');
      clearAllBtn.textContent = t.clearAllConfirm;
      armTimer = setTimeout(resetArmed, 3000);
      return;
    }
    resetArmed();
    if (wb.clearConditionalFormats(sheet)) {
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
