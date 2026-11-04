import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { setActive } from './fixtures.js';

describe('Data Validation dialog tabs', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    setActive(store, 0, 0);
    attachFormatDialog({ host, store }).open('more', { mode: 'dataValidation' });
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  const tabs = (): HTMLButtonElement[] =>
    Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__vtabs [role="tab"]'));
  const panes = (): HTMLElement[] =>
    Array.from(document.querySelectorAll<HTMLElement>('.fc-fmtdlg__vpane'));
  const kindSelect = (): HTMLSelectElement => {
    const select = document
      .querySelector<HTMLElement>('.fc-fmtdlg__vpane')
      ?.querySelector<HTMLSelectElement>('select');
    if (!select) throw new Error('missing kind select');
    return select;
  };

  it('lays the editor out as four tabs with the settings tab selected', () => {
    expect(tabs().map((tab) => tab.textContent)).toEqual([
      '設定',
      '入力時メッセージ',
      'エラー メッセージ',
      '日本語入力',
    ]);
    expect(tabs().map((tab) => tab.getAttribute('aria-selected'))).toEqual([
      'true',
      'false',
      'false',
      'false',
    ]);
    expect(panes().map((pane) => pane.hidden)).toEqual([false, true, true, true]);
  });

  it('defaults Allow to Any value and labels the data fields', () => {
    expect(kindSelect().value).toBe('none');
    expect(kindSelect().selectedOptions[0]?.textContent).toBe('すべての値');
    expect(kindSelect().getAttribute('aria-label')).toBe('許可:');
  });

  it('switches panes when a tab is pressed', () => {
    tabs()[2]?.click();
    expect(panes().map((pane) => pane.hidden)).toEqual([true, true, false, true]);
    expect(tabs()[2]?.getAttribute('aria-selected')).toBe('true');
    expect(tabs()[0]?.getAttribute('aria-selected')).toBe('false');
  });

  it('keeps the input message and error alert fields on their own tabs', () => {
    const [, input, error] = panes();
    expect(input?.textContent).toContain('セルを選択したときに入力時メッセージを表示する');
    expect(error?.textContent).toContain('無効なデータが入力されたらエラー メッセージを表示する');
    expect(error?.textContent).toContain('スタイル:');
  });

  it('offers only the no-control mode on the Japanese input tab', () => {
    const ime = panes()[3]?.querySelector<HTMLSelectElement>('select');
    expect(ime?.getAttribute('aria-label')).toBe('日本語入力:');
    expect(Array.from(ime?.options ?? [], (o) => o.textContent)).toEqual(['コントロールなし']);
  });

  it('Clear All resets the rule to Any value', () => {
    const select = kindSelect();
    select.value = 'list';
    select.dispatchEvent(new Event('change', { bubbles: true }));
    const clear = Array.from(document.querySelectorAll<HTMLButtonElement>('button')).find(
      (b) => b.textContent === 'すべてクリア',
    );
    if (!clear) throw new Error('missing Clear All');
    clear.click();
    expect(select.value).toBe('none');
  });
});
