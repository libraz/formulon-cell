import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { hydrateConditionalRulesFromEngine } from '../../../../src/engine/cf-sync.js';
import { conditionalRuleToEngineInput, inputWithDxf } from '../../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { setRange } from './fixtures.js';

const REFUSED = '=SUM()';
const A1 = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 } as const;

describe('conditional dialog with an engine that refuses the rule', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
  });

  afterEach(() => {
    wb.dispose();
    document.body.innerHTML = '';
  });

  const clickAdd = (): void => {
    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  };
  const alertText = (): string | null => {
    const row = document.querySelector<HTMLElement>('.fc-conddlg [role="alert"]');
    return row && !row.hidden ? row.textContent : null;
  };

  it('keeps a refused new rule out of the store and history and keeps the dialog open', () => {
    setRange(store, 0, 0, 0, 0);
    const history = new History();
    const handle = attachConditionalDialog({ strings: en, host, store, history, getWb: () => wb });
    handle.open({ mode: 'new', kind: 'formula' });
    const input = document
      .querySelector<HTMLInputElement>('[data-range-picker="conditional-format-formula"]')
      ?.closest('.fc-range-picker')
      ?.querySelector('input');
    if (!input) throw new Error('missing formula input');
    input.value = REFUSED;
    clickAdd();

    expect(store.getState().conditional.rules).toHaveLength(0);
    expect(history.canUndo()).toBe(false);
    expect(wb.getConditionalFormats(0)).toHaveLength(0);
    expect(document.querySelector<HTMLElement>('.fc-conddlg')?.hidden).toBe(false);
    expect(alertText()).toBe(en.commandRefusal.engineFailed);
    handle.detach();
  });

  it('takes an accepted new rule and closes', () => {
    setRange(store, 0, 0, 0, 0);
    const history = new History();
    const handle = attachConditionalDialog({ strings: en, host, store, history, getWb: () => wb });
    handle.open({ mode: 'new', kind: 'formula' });
    const input = document
      .querySelector<HTMLInputElement>('[data-range-picker="conditional-format-formula"]')
      ?.closest('.fc-range-picker')
      ?.querySelector('input');
    if (!input) throw new Error('missing formula input');
    input.value = '=A1>1';
    clickAdd();

    expect(store.getState().conditional.rules).toHaveLength(1);
    expect(history.canUndo()).toBe(true);
    expect(wb.getConditionalFormats(0)).toHaveLength(0);
    expect(document.querySelector<HTMLElement>('.fc-conddlg')?.hidden).toBe(true);
    handle.detach();
  });

  it('keeps a loaded rule as it was when the engine refuses its edit', () => {
    const imported = {
      kind: 'cell-value',
      range: A1,
      op: '>',
      a: 10,
      apply: { fill: '#ffc7ce' },
    } as const;
    const input = conditionalRuleToEngineInput(imported);
    if (!input) throw new Error('untranslatable rule');
    wb.addConditionalFormat(0, inputWithDxf(wb, input, imported));
    hydrateConditionalRulesFromEngine(wb, store, 0);
    const before = wb.getConditionalFormats(0)[0];

    const handle = attachConditionalDialog({ strings: en, host, store, getWb: () => wb });
    handle.open({ mode: 'edit', editIndex: 0 });
    const valueInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__sub input[type="text"]',
    );
    if (!valueInput) throw new Error('missing value input');
    valueInput.value = REFUSED;
    clickAdd();

    expect(wb.getConditionalFormats(0)).toHaveLength(1);
    expect(wb.getConditionalFormats(0)[0]).toMatchObject({
      id: before?.id,
      formula1: before?.formula1,
    });
    expect(store.getState().conditional.rules[0]).toMatchObject({ a: 10 });
    expect(document.querySelector<HTMLElement>('.fc-conddlg')?.hidden).toBe(false);
    expect(alertText()).toBe(en.commandRefusal.engineFailed);
    handle.detach();
  });

  describe('the engine check leaves no trace', () => {
    const addFormulaRule = (formula: string): void => {
      setRange(store, 0, 0, 0, 0);
      const handle = attachConditionalDialog({ strings: en, host, store, getWb: () => wb });
      handle.open({ mode: 'new', kind: 'formula' });
      const input = document
        .querySelector('[data-range-picker="conditional-format-formula"]')
        ?.closest('.fc-range-picker')
        ?.querySelector('input');
      if (!input) throw new Error('missing formula input');
      input.value = formula;
      clickAdd();
      handle.detach();
    };
    const seed = (): void => {
      for (const priority of [1, 2]) {
        wb.addConditionalFormat(0, {
          ranges: [{ firstRow: 5, firstCol: 0, lastRow: 5, lastCol: 0 }],
          type: 17,
          priority,
        });
      }
    };

    it('adds no dxf and no rule for an accepted rule that carries a fill', () => {
      seed();
      const before = wb.getConditionalFormats(0);
      const dxfs = wb.dxfCount();
      addFormulaRule('=A1>1');
      expect(wb.dxfCount()).toBe(dxfs);
      expect(wb.getConditionalFormats(0)).toEqual(before);
    });

    it('adds no dxf for a refused rule', () => {
      seed();
      const dxfs = wb.dxfCount();
      addFormulaRule(REFUSED);
      expect(wb.dxfCount()).toBe(dxfs);
    });
  });
});
