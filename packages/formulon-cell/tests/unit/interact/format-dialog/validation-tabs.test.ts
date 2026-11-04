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

  it('keeps the message tabs filled while Allow is Any value', () => {
    const pane = (i: number): HTMLElement => panes()[i] as HTMLElement;
    const shown = (el: Element): boolean =>
      !(el as HTMLElement).closest('[hidden]:not(.fc-fmtdlg__vpane)');
    tabs()[1]?.click();
    const input = pane(1);
    const inputChecks = [...input.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')];
    expect(inputChecks.some((c) => shown(c) && c.checked)).toBe(true);
    expect([...input.querySelectorAll('input[type="text"], textarea')].filter(shown).length).toBe(
      2,
    );
    tabs()[2]?.click();
    const error = pane(2);
    expect([...error.querySelectorAll('select')].filter(shown).length).toBe(1);
    expect([...error.querySelectorAll('input[type="text"], textarea')].filter(shown).length).toBe(
      2,
    );
  });

  it('captions the Mac message tabs like the desktop app', () => {
    document.body.innerHTML = '';
    const macHost = document.createElement('div');
    macHost.className = 'fc-host';
    macHost.dataset.fcPlatform = 'mac';
    document.body.appendChild(macHost);
    const store = createSpreadsheetStore();
    setActive(store, 0, 0);
    attachFormatDialog({ host: macHost, store }).open('more', { mode: 'dataValidation' });
    const caption = (pane: string): string | null | undefined =>
      document.querySelector(`.fc-fmtdlg__vpane[data-fc-pane="${pane}"] .fc-fmtdlg__vmsg-caption`)
        ?.textContent;
    expect(caption('input')).toBe('セルを選択したときに、この入力時メッセージを表示:');
    expect(caption('error')).toBe('無効なデータが入力されたときに表示するエラー メッセージ:');
    const icon = (): string | undefined =>
      document.querySelector<HTMLElement>('.fc-fmtdlg__valert-icon')?.dataset.style;
    expect(icon()).toBe('stop');
    const style = document.querySelector<HTMLSelectElement>(
      '.fc-fmtdlg__vpane[data-fc-pane="error"] select',
    );
    if (!style) throw new Error('missing style select');
    style.value = 'warning';
    style.dispatchEvent(new Event('change', { bubbles: true }));
    expect(icon()).toBe('warning');
    expect(document.querySelector('.fc-fmtdlg__valert-icon svg')).not.toBeNull();
  });
});
