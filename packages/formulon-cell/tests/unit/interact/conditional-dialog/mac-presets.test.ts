import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { en } from '../../../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { applyConditionalMenuAction } from '../../../../src/toolbar/ribbon/conditional-menu-action.js';
import { setRange } from './fixtures.js';

// Excel for Mac: every Highlight Cells / Top-Bottom preset opens the full "New Formatting Rule"
// dialog pre-set to the rule, with empty value fields and no separate colour rows.
describe('Mac conditional formatting presets', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setRange(store, 1, 1, 4, 3);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('opens a preset with empty value fields and without the range and colour rows', () => {
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'new', kind: 'cell-value', cellValueOp: '>' });
    const overlay = document.querySelector<HTMLElement>('.fc-conddlg');
    expect(overlay?.classList.contains('fc-conddlg--mac')).toBe(true);
    const values = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-conddlg__sub input[type="text"]'),
    ).filter((input) => input.closest<HTMLElement>('.fc-conddlg__sub')?.hidden === false);
    expect(values.every((input) => input.value === '')).toBe(true);
    expect(document.querySelector<HTMLElement>('.fc-conddlg__range-row')?.hidden).toBe(true);
    expect(
      Array.from(document.querySelectorAll<HTMLElement>('.fc-conddlg__apply-box')).every(
        (box) => box.hidden,
      ),
    ).toBe(true);
  });

  it('keeps the desktop dialog unchanged off the Mac profile', () => {
    host.dataset.fcPlatform = 'default';
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'new', kind: 'cell-value', cellValueOp: '>' });
    expect(document.querySelector('.fc-conddlg')?.classList.contains('fc-conddlg--mac')).toBe(
      false,
    );
    expect(document.querySelector<HTMLElement>('.fc-conddlg__range-row')?.hidden).toBe(false);
  });

  it.each([
    ['cell-gt', { kind: 'cell-value', cellValueOp: '>' }],
    ['cell-lt', { kind: 'cell-value', cellValueOp: '<' }],
    ['cell-eq', { kind: 'cell-value', cellValueOp: '=' }],
    ['cell-between', { kind: 'cell-value', cellValueOp: 'between' }],
    ['text-contains', { kind: 'text-contains' }],
    ['date-occurring', { kind: 'date-occurring', datePeriod: 'yesterday' }],
    ['duplicates', { kind: 'duplicates' }],
    ['top10', { kind: 'top-bottom', topBottomMode: 'top', topBottomPercent: false }],
    ['top10-percent', { kind: 'top-bottom', topBottomMode: 'top', topBottomPercent: true }],
    ['bottom10', { kind: 'top-bottom', topBottomMode: 'bottom', topBottomPercent: false }],
    ['bottom10-percent', { kind: 'top-bottom', topBottomMode: 'bottom', topBottomPercent: true }],
    ['above-avg', { kind: 'average', averageMode: 'above' }],
    ['below-avg', { kind: 'average', averageMode: 'below' }],
  ])('routes the %s preset to the full dialog', async (action, expected) => {
    const openConditionalDialog = vi.fn();
    const promptCfNumber = vi.fn();
    await applyConditionalMenuAction(
      {
        inst: { host, openConditionalDialog } as never,
        ribbonLang: 'ja',
        range: { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 3 },
        cfFill: { fill: '#ffc7ce', color: '#9c0006' },
        promptCfNumber,
        promptCfText: vi.fn(),
        showChoiceDialog: vi.fn(),
        showMessage: vi.fn(),
        refreshWorkbookCells: vi.fn(),
        addConditionalRuleFromRibbon: vi.fn(),
      },
      action,
    );
    expect(openConditionalDialog).toHaveBeenCalledWith({ mode: 'new', ...expected });
    expect(promptCfNumber).not.toHaveBeenCalled();
  });
});
