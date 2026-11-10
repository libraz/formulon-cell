import { CfRuleType } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachCfRulesDialog } from '../../../src/interact/cf-rules-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type Rule = ReturnType<WorkbookHandle['getConditionalFormats']>[number];

const rule = (overrides: Partial<Rule> = {}): Rule =>
  ({
    id: 'r1',
    type: CfRuleType.CellIs,
    priority: 1,
    stopIfTrue: false,
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
    ...overrides,
  }) as Rule;

const fakeWb = (rules: Rule[]): { wb: WorkbookHandle; removed: number[] } => {
  let snapshot = [...rules];
  const removed: number[] = [];
  const wb = {
    getConditionalFormats: () => snapshot,
    removeConditionalFormatAt: (_sheet: number, idx: number): boolean => {
      removed.push(idx);
      snapshot = snapshot.filter((_, i) => i !== idx);
      return true;
    },
    clearConditionalFormats: (): boolean => true,
  } as unknown as WorkbookHandle;
  return { wb, removed };
};

const q = <T extends HTMLElement>(sel: string): T => {
  const el = document.querySelector<T>(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el;
};
const labels = (sel: string): (string | null)[] =>
  Array.from(document.querySelectorAll<HTMLElement>(sel)).map((el) => el.textContent);

describe('attachCfRulesDialog on the Mac profile', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  const open = (rules: Rule[], extra: Partial<Parameters<typeof attachCfRulesDialog>[0]> = {}) => {
    const fake = fakeWb(rules);
    const handle = attachCfRulesDialog({
      host,
      getWb: () => fake.wb,
      getActiveSheet: () => 0,
      strings: ja,
      ...extra,
    });
    handle.open();
    return { ...fake, handle };
  };

  it('opens on the current selection scope without the read-only note', () => {
    const { handle } = open([rule()]);
    expect(q<HTMLSelectElement>('.fc-cfrulesdlg__scope-select').value).toBe('selection');
    expect(document.querySelector('.fc-cfrulesdlg__note')).toBeNull();
    handle.detach();
  });

  it('lists the four rule columns without an actions column', () => {
    const { handle } = open([rule()]);
    expect(labels('.fc-cfrulesdlg__table thead th')).toEqual([
      'ルール (表示順で適用)',
      ja.cfRulesDialog.headerFormat,
      ja.cfRulesDialog.headerAppliesTo,
      ja.cfRulesDialog.headerStopIfTrue,
    ]);
    handle.detach();
  });

  it('puts + / - / edit / duplicate in a bottom bar and the order arrows beside the label', () => {
    const { handle } = open([rule()]);
    expect(document.querySelector('.fc-cfrulesdlg__commandbar')).toBeNull();
    expect(labels('.fc-cfrulesdlg__bottombar button')).toEqual([
      '+',
      '−',
      ja.cfRulesDialog.editRule,
      'ルールの重複',
    ]);
    expect(document.querySelectorAll('.fc-cfrulesdlg__order button')).toHaveLength(2);
    handle.detach();
  });

  it('ends with Cancel and OK instead of Clear All and Close', () => {
    const { handle } = open([rule()]);
    expect(labels('.fc-cfrulesdlg__footer > button')).toEqual([
      ja.formatDialog.cancel,
      ja.formatDialog.ok,
    ]);
    handle.detach();
  });

  it('shows the selected rule range in the Applies to row', () => {
    const { handle } = open([rule()]);
    expect(q<HTMLInputElement>('.fc-cfrulesdlg__appliesto input').value).toBe('A1:A5');
    handle.detach();
  });

  it('keeps a removed loaded rule until OK and drops the change on Cancel', () => {
    const cancelled = open([rule({ id: 'a' }), rule({ id: 'b', priority: 2 })]);
    q<HTMLButtonElement>('.fc-cfrulesdlg__bottombar button:nth-child(2)').click();
    expect(document.querySelectorAll('.fc-cfrulesdlg__table tbody tr')).toHaveLength(1);
    expect(cancelled.removed).toEqual([]);
    q<HTMLButtonElement>('.fc-cfrulesdlg__footer > button:not(.fc-fmtdlg__btn--primary)').click();
    expect(cancelled.removed).toEqual([]);
    cancelled.handle.detach();
    document.body.innerHTML = '';
    host = document.createElement('div');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);

    const confirmed = open([rule({ id: 'a' }), rule({ id: 'b', priority: 2 })]);
    q<HTMLButtonElement>('.fc-cfrulesdlg__bottombar button:nth-child(2)').click();
    q<HTMLButtonElement>('.fc-cfrulesdlg__footer .fc-fmtdlg__btn--primary').click();
    expect(confirmed.removed).toEqual([0]);
    confirmed.handle.detach();
  });

  it('restores removed session rules on Cancel', () => {
    const store = createSpreadsheetStore();
    mutators.addConditionalRule(store, {
      kind: 'duplicates',
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      apply: { fill: '#ff0000' },
    });
    const { handle } = open([], { store });
    q<HTMLButtonElement>('.fc-cfrulesdlg__bottombar button:nth-child(2)').click();
    expect(store.getState().conditional.rules).toHaveLength(0);
    q<HTMLButtonElement>('.fc-cfrulesdlg__footer > button:not(.fc-fmtdlg__btn--primary)').click();
    expect(store.getState().conditional.rules).toHaveLength(1);
    handle.detach();
  });

  it('uses the desktop wording and keeps the column headers over an empty list', () => {
    const { handle } = open([]);
    expect(document.querySelector('.fc-cfrulesdlg .fc-fmtdlg__header')?.textContent).toBe(
      'ルールの管理',
    );
    expect(labels('.fc-cfrulesdlg__order-label')).toEqual(['ルールの順序を変更:']);
    expect(labels('.fc-cfrulesdlg__table thead th')).toEqual([
      'ルール (表示順で適用)',
      '書式',
      '適用先',
      '条件を満たす場合は停止',
    ]);
    expect(document.querySelectorAll('.fc-cfrulesdlg__table tbody tr')).toHaveLength(0);
    expect(
      document.querySelector('.fc-cfrulesdlg__appliesto')?.firstElementChild?.textContent,
    ).toBe('適用対象:');
    // The + / - group shares the footer row with Cancel / OK.
    expect(
      document.querySelector('.fc-cfrulesdlg__footer .fc-cfrulesdlg__bottombar'),
    ).not.toBeNull();
    handle.detach();
  });

  it('renders the Applies to editor disabled until a rule is selected', () => {
    const empty = open([]);
    const row = q<HTMLElement>('.fc-cfrulesdlg__appliesto');
    const input = q<HTMLInputElement>('.fc-cfrulesdlg__appliesto input');
    expect(input.disabled).toBe(true);
    const buttons = [...row.querySelectorAll<HTMLButtonElement>('button')];
    expect(buttons.map((b) => b.className.match(/--(cancel|commit)\b/)?.[1] ?? 'picker')).toEqual([
      'cancel',
      'commit',
      'picker',
    ]);
    expect(buttons.every((b) => b.disabled)).toBe(true);
    empty.handle.detach();
    document.body.innerHTML = '';
    host = document.createElement('div');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);

    const { handle } = open([rule()]);
    // A loaded rule is editable only when the engine can rewrite it.
    expect(q<HTMLInputElement>('.fc-cfrulesdlg__appliesto input').disabled).toBe(true);
    expect(
      [...document.querySelectorAll<HTMLButtonElement>('.fc-cfrulesdlg__appliesto button')].every(
        (b) => b.disabled,
      ),
    ).toBe(true);
    handle.detach();
  });

  it('sizes the four rule columns with a colgroup', () => {
    const { handle } = open([]);
    expect(document.querySelectorAll('.fc-cfrulesdlg__table colgroup col')).toHaveLength(4);
    handle.detach();
  });

  describe('Applies to editing', () => {
    const setSelection = (
      store: ReturnType<typeof createSpreadsheetStore>,
      r0: number,
      c0: number,
      r1: number,
      c1: number,
    ): void => {
      store.setState((s) => ({
        ...s,
        selection: {
          active: { sheet: 0, row: r0, col: c0 },
          anchor: { sheet: 0, row: r0, col: c0 },
          range: { sheet: 0, r0, c0, r1, c1 },
        },
      }));
    };
    const sessionRule = {
      kind: 'blanks' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      apply: {},
    };

    it('takes a grid pick into the focused field and stages it for OK', () => {
      const store = createSpreadsheetStore();
      mutators.addConditionalRule(store, sessionRule);
      const { handle } = open([], { store });
      const input = q<HTMLInputElement>('.fc-cfrulesdlg__appliesto input');
      expect(input.disabled).toBe(false);
      input.focus();
      setSelection(store, 1, 1, 3, 2);
      expect(input.value).toBe('$B$2:$C$4');
      expect(q<HTMLElement>('.fc-cfrulesdlg').hidden).toBe(false);
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      expect(store.getState().conditional.rules[0]?.range).toMatchObject({
        r0: 1,
        c0: 1,
        r1: 3,
        c1: 2,
      });
      expect(labels('.fc-cfrulesdlg__cell-range')).toEqual(['B2:C4']);
      handle.detach();
    });

    it('drops a picked range on Cancel', () => {
      const store = createSpreadsheetStore();
      mutators.addConditionalRule(store, sessionRule);
      const { handle } = open([], { store });
      const input = q<HTMLInputElement>('.fc-cfrulesdlg__appliesto input');
      input.focus();
      setSelection(store, 1, 1, 3, 2);
      input.dispatchEvent(new Event('change', { bubbles: true }));
      [...document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button')]
        .find((b) => b.textContent === ja.formatDialog.cancel)
        ?.click();
      expect(store.getState().conditional.rules[0]?.range).toMatchObject({ r0: 0, r1: 4 });
      handle.detach();
    });
  });
});
