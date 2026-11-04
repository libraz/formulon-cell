import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachCfRulesDialog } from '../../../src/interact/cf-rules-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type Rule = ReturnType<WorkbookHandle['getConditionalFormats']>[number];

const rule = (overrides: Partial<Rule> = {}): Rule => ({
  id: 'r1',
  type: 1,
  priority: 1,
  stopIfTrue: false,
  sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
  ...overrides,
});

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
      defaultStrings.cfRulesDialog.headerFormat,
      defaultStrings.cfRulesDialog.headerAppliesTo,
      defaultStrings.cfRulesDialog.headerStopIfTrue,
    ]);
    handle.detach();
  });

  it('puts + / - / edit / duplicate in a bottom bar and the order arrows beside the label', () => {
    const { handle } = open([rule()]);
    expect(document.querySelector('.fc-cfrulesdlg__commandbar')).toBeNull();
    expect(labels('.fc-cfrulesdlg__bottombar button')).toEqual([
      '+',
      '−',
      defaultStrings.cfRulesDialog.editRule,
      'ルールの重複',
    ]);
    expect(document.querySelectorAll('.fc-cfrulesdlg__order button')).toHaveLength(2);
    handle.detach();
  });

  it('ends with Cancel and OK instead of Clear All and Close', () => {
    const { handle } = open([rule()]);
    expect(labels('.fc-cfrulesdlg__footer > button')).toEqual([
      defaultStrings.formatDialog.cancel,
      defaultStrings.formatDialog.ok,
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
    expect(labels('.fc-cfrulesdlg__appliesto > span')).toEqual(['適用対象:']);
    // The + / - group shares the footer row with Cancel / OK.
    expect(
      document.querySelector('.fc-cfrulesdlg__footer .fc-cfrulesdlg__bottombar'),
    ).not.toBeNull();
    handle.detach();
  });
});
