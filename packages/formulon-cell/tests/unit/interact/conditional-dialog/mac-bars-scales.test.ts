import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { ja } from '../../../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { setRange } from './fixtures.js';

// Excel for Mac (ja): the Data Bar and Color Scale editors of New Formatting Rule.
describe('Mac data bar and color scale editors', () => {
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

  const open = (kind: 'data-bar' | 'color-scale') => {
    const handle = attachConditionalDialog({ host, store, strings: ja });
    handle.open({ mode: 'new', kind });
    return handle;
  };
  const visibleText = (selector: string): string[] =>
    Array.from(document.querySelectorAll<HTMLElement>(selector))
      .filter((el) => !el.closest('[hidden]') && !el.hidden)
      .map((el) => el.textContent ?? '');

  it('splits fill and border into positive and negative values under their own headings', () => {
    open('data-bar');
    expect(visibleText('.fc-conddlg__mac-heading')).toEqual(['フィル', '罫線']);
    const labels = visibleText('.fc-conddlg__bar-sign');
    expect(labels).toEqual(['正の値:', '負の値:', '正の値:', '負の値:']);
  });

  it('shows "bar only" as the inverse of the model\'s show-value flag', () => {
    open('data-bar');
    const barOnly = document.querySelector<HTMLInputElement>('input[data-cf-bar-only]');
    const showValue = document.querySelector<HTMLInputElement>('input[data-cf-bar-show-value]');
    if (!barOnly || !showValue) throw new Error('missing bar-only controls');
    expect(barOnly.checked).toBe(false);
    expect(showValue.checked).toBe(true);
    expect(barOnly.closest<HTMLElement>('.fc-fmtdlg__check')?.textContent).toBe(
      'データ バーのみ表示',
    );
    barOnly.click();
    expect(showValue.checked).toBe(false);
    barOnly.click();
    expect(showValue.checked).toBe(true);
  });

  it('keeps an automatic bound as a disabled field with its name in parentheses', () => {
    open('color-scale');
    const shown = (placeholder: string): HTMLInputElement[] =>
      Array.from(
        document.querySelectorAll<HTMLInputElement>(`input[placeholder="${placeholder}"]`),
      ).filter((input) => !input.closest('[hidden]'));
    expect(shown('(最大値)')).toHaveLength(1);
    const [minField] = shown('(最小値)');
    expect(shown('(最小値)')).toHaveLength(1);
    expect(minField?.disabled).toBe(true);
    expect(minField?.hidden).toBe(false);
  });

  it('leaves the desktop dialog as it was', () => {
    document.body.innerHTML = '';
    host = document.createElement('div');
    host.dataset.fcPlatform = 'default';
    document.body.appendChild(host);
    open('data-bar');
    expect(visibleText('.fc-conddlg__mac-heading')).toEqual([]);
    expect(
      document.querySelector<HTMLElement>('input[data-cf-bar-only]')?.closest('[hidden]'),
    ).not.toBeNull();
  });
});
