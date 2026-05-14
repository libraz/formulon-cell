import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachCfRulesDialog } from '../../../src/interact/cf-rules-dialog.js';

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
    expect(rows[0]?.textContent).toContain('cellIs');
    expect(rows[0]?.textContent).toContain('A1:A5');
    expect(rows[1]?.textContent).toContain('duplicateValues');
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
