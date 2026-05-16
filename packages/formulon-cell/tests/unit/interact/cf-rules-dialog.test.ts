import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachCfRulesDialog } from '../../../src/interact/cf-rules-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type Rule = ReturnType<WorkbookHandle['getConditionalFormats']>[number];

const rule = (overrides: Partial<Rule> = {}): Rule => ({
  id: 'r1',
  type: 1, // cellIs
  priority: 1,
  stopIfTrue: false,
  sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
  ...overrides,
});

interface FakeWbHandle {
  wb: WorkbookHandle;
  removed: number[];
  /** Wrap counter in an object so test code reads the live mutated value
   *  rather than a snapshot taken at fakeWb() return time. */
  state: { cleared: number };
}

const fakeWb = (rules: Rule[]): FakeWbHandle => {
  let snapshot = [...rules];
  const removed: number[] = [];
  const state = { cleared: 0 };
  const wb = {
    getConditionalFormats: () => snapshot,
    removeConditionalFormatAt: (_sheet: number, idx: number): boolean => {
      removed.push(idx);
      snapshot = snapshot.filter((_, i) => i !== idx);
      return true;
    },
    clearConditionalFormats: (): boolean => {
      state.cleared += 1;
      snapshot = [];
      return true;
    },
  } as unknown as WorkbookHandle;
  return { wb, removed, state };
};

describe('attachCfRulesDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('renders an empty-state when the engine reports no rules', () => {
    const { wb } = fakeWb([]);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
    });
    handle.open();
    const empty = document.querySelector<HTMLElement>('.fc-cfrulesdlg__empty');
    expect(empty?.hidden).toBe(false);
    const clearAll = document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__clearall');
    expect(clearAll?.disabled).toBe(true);
    handle.detach();
  });

  it('renders one row per rule with priority/type/range columns', () => {
    const { wb } = fakeWb([
      rule({ id: 'a', priority: 1, type: 1 }),
      rule({
        id: 'b',
        priority: 2,
        type: 16, // duplicateValues
        sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
      }),
    ]);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
    });
    handle.open();
    const rows = document.querySelectorAll<HTMLTableRowElement>('.fc-cfrulesdlg__table tbody tr');
    expect(rows.length).toBe(2);
    expect(rows[0]?.textContent).toContain(defaultStrings.cfRulesDialog.ruleCellIs);
    expect(rows[0]?.textContent).toContain('A1:A5');
    expect(rows[0]?.tabIndex).toBe(0);
    expect(rows[0]?.getAttribute('aria-selected')).toBe('true');
    expect(rows[1]?.tabIndex).toBe(-1);
    expect(rows[1]?.textContent).toContain(defaultStrings.cfRulesDialog.ruleDuplicateValues);
    handle.detach();
  });

  it('renders localized rule type labels', () => {
    const { wb } = fakeWb([rule({ type: 3 })]);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      strings: {
        ...defaultStrings,
        cfRulesDialog: {
          ...defaultStrings.cfRulesDialog,
          ruleDataBar: 'データ バー',
        },
      },
    });
    handle.open();
    const row = document.querySelector<HTMLTableRowElement>('.fc-cfrulesdlg__table tbody tr');
    expect(row?.textContent).toContain('データ バー');
    handle.detach();
  });

  it('supports Excel-style row selection keys and Delete removal', () => {
    const { wb, removed } = fakeWb([
      rule({ id: 'a', priority: 1 }),
      rule({ id: 'b', priority: 2 }),
      rule({ id: 'c', priority: 3 }),
    ]);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
    });
    handle.open();
    const rows = (): HTMLTableRowElement[] =>
      Array.from(document.querySelectorAll<HTMLTableRowElement>('.fc-cfrulesdlg__table tbody tr'));

    rows()[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(rows()[1]?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(rows()[1]);

    rows()[1]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(rows()[2]?.getAttribute('aria-selected')).toBe('true');

    rows()[2]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(rows()[0]?.getAttribute('aria-selected')).toBe('true');

    rows()[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));
    expect(removed).toEqual([0]);
    expect(rows()).toHaveLength(2);
    expect(rows()[0]?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('clicking the per-row remove button calls removeConditionalFormatAt and rerenders', () => {
    const { wb, removed } = fakeWb([
      rule({ id: 'a', priority: 1 }),
      rule({ id: 'b', priority: 2 }),
    ]);
    let changed = 0;
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      onChanged: () => {
        changed += 1;
      },
    });
    handle.open();
    const buttons = document.querySelectorAll<HTMLButtonElement>('.fc-cfrulesdlg__remove');
    expect(buttons.length).toBe(2);
    buttons[0]?.click();
    expect(removed).toEqual([0]);
    expect(changed).toBe(1);
    const after = document.querySelectorAll('.fc-cfrulesdlg__remove');
    expect(after.length).toBe(1);
    handle.detach();
  });

  it('lists session conditional rules and removes them through undoable store history', () => {
    const { wb } = fakeWb([]);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 },
      color: '#4472c4',
      gradient: true,
    });
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      store,
      history,
    });
    handle.open();
    const row = document.querySelector<HTMLTableRowElement>('.fc-cfrulesdlg__table tbody tr');
    expect(row?.textContent).toContain(defaultStrings.cfRulesDialog.ruleDataBar);
    expect(row?.textContent).toContain('B2:B4');

    document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__remove')?.click();
    expect(store.getState().conditional.rules).toHaveLength(0);
    expect(history.undo()).toBe(true);
    expect(store.getState().conditional.rules).toHaveLength(1);
    handle.detach();
  });

  it('duplicates session conditional rules through undoable store history', () => {
    const { wb } = fakeWb([]);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 },
      stops: ['#f8696b', '#63be7b'],
      thresholds: [{ kind: 'min' }, { kind: 'max' }],
    });
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      store,
      history,
    });
    handle.open();

    document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__duplicate')?.click();

    expect(store.getState().conditional.rules).toHaveLength(2);
    expect(store.getState().conditional.rules[1]).toEqual(store.getState().conditional.rules[0]);
    expect(store.getState().conditional.rules[1]).not.toBe(store.getState().conditional.rules[0]);
    expect(history.undo()).toBe(true);
    expect(store.getState().conditional.rules).toHaveLength(1);
    expect(history.redo()).toBe(true);
    expect(store.getState().conditional.rules).toHaveLength(2);
    handle.detach();
  });

  it('moves session conditional rules up and down through undoable store history', () => {
    const { wb } = fakeWb([]);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      color: '#4472c4',
    });
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 },
      stops: ['#f8696b', '#63be7b'],
    });
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      store,
      history,
    });
    handle.open();

    document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__move-down')?.click();
    expect(store.getState().conditional.rules.map((rule) => rule.kind)).toEqual([
      'color-scale',
      'data-bar',
    ]);
    expect(history.undo()).toBe(true);
    expect(store.getState().conditional.rules.map((rule) => rule.kind)).toEqual([
      'data-bar',
      'color-scale',
    ]);
    expect(history.redo()).toBe(true);
    expect(store.getState().conditional.rules.map((rule) => rule.kind)).toEqual([
      'color-scale',
      'data-bar',
    ]);

    document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__move-up:not(:disabled)')?.click();
    expect(store.getState().conditional.rules.map((rule) => rule.kind)).toEqual([
      'data-bar',
      'color-scale',
    ]);
    handle.detach();
  });

  it('clearAll clears both engine and session rules', () => {
    const { wb, state } = fakeWb([rule()]);
    const store = createSpreadsheetStore();
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      stops: ['#f8696b', '#63be7b'],
    });
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      store,
    });
    handle.open();
    const clearAll = document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__clearall');
    clearAll?.click();
    clearAll?.click();
    expect(state.cleared).toBe(1);
    expect(store.getState().conditional.rules).toHaveLength(0);
    handle.detach();
  });

  it('clearAll requires two clicks (arm-then-confirm); first click flips the label', () => {
    const { wb, state } = fakeWb([rule()]);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
    });
    handle.open();
    const clearAll = document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__clearall');
    const initialLabel = clearAll?.textContent;
    clearAll?.click();
    expect(state.cleared).toBe(0);
    expect(clearAll?.classList.contains('fc-cfrulesdlg__clearall--armed')).toBe(true);
    expect(clearAll?.textContent).not.toBe(initialLabel);
    clearAll?.click();
    expect(state.cleared).toBe(1);
    handle.detach();
  });

  it('New Rule closes Manage Rules and opens the companion rule dialog callback', () => {
    const { wb } = fakeWb([]);
    let openedNewRule = 0;
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      onNewRule: () => {
        openedNewRule += 1;
      },
    });
    handle.open();

    document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__new')?.click();

    expect(openedNewRule).toBe(1);
    expect(document.querySelector<HTMLElement>('.fc-cfrulesdlg')?.hidden).toBe(true);
    handle.detach();
  });

  it('Escape closes the dialog and disarms a pending clearAll', () => {
    const { wb } = fakeWb([rule()]);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
    });
    handle.open();
    const clearAll = document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__clearall');
    clearAll?.click(); // arm
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    const dialog = document.querySelector<HTMLElement>('.fc-cfrulesdlg');
    expect(dialog?.hidden).toBe(true);
    expect(clearAll?.classList.contains('fc-cfrulesdlg__clearall--armed')).toBe(false);
    handle.detach();
  });

  it('detach removes the dialog node from the DOM', () => {
    const { wb } = fakeWb([]);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
    });
    handle.detach();
    expect(document.querySelector('.fc-cfrulesdlg')).toBeNull();
  });
});
