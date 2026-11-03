import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachFindReplace } from '../../../src/interact/find-replace.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedText = (store: SpreadsheetStore, wb: WorkbookHandle, value: string): void => {
  const addr = { sheet: 0, row: 0, col: 0 };
  wb.setText(addr, value);
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey(addr), { value: { kind: 'text', value }, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

const mount = async (platform: 'mac' | 'default', text?: string) => {
  const fcHost = document.createElement('div');
  fcHost.className = 'fc-host';
  fcHost.dataset.fcPlatform = platform;
  const host = document.createElement('div');
  host.tabIndex = -1;
  fcHost.appendChild(host);
  document.body.appendChild(fcHost);
  const store = createSpreadsheetStore();
  const wb = await newWb();
  if (text) seedText(store, wb, text);
  const handle = attachFindReplace({ host, store, wb, strings: ja });
  return { host, store, handle };
};

const q = <T extends Element>(host: HTMLElement, sel: string): T => {
  const el = host.querySelector(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el as T;
};

afterEach(() => {
  document.body.innerHTML = '';
});

describe('find and replace, Mac Japanese presentation', () => {
  let ctx: Awaited<ReturnType<typeof mount>>;
  beforeEach(async () => {
    ctx = await mount('mac');
  });

  it('uses the Mac app wording and defaults', () => {
    ctx.handle.open('find');
    const optionsBtn = q<HTMLButtonElement>(ctx.host, '.fc-find__options-btn');
    expect(optionsBtn.textContent).toBe('オプション');
    optionsBtn.click();
    expect(optionsBtn.textContent).toBe('オプション');
    const labels = [...ctx.host.querySelectorAll('.fc-find__field label')].map(
      (l) => l.textContent,
    );
    expect(labels).toEqual(
      expect.arrayContaining(['検索する文字列:', '検索場所:', '検索:', '検索先:']),
    );
    const checks = [...ctx.host.querySelectorAll('.fc-find__check span')].map((s) => s.textContent);
    expect(checks).toEqual([
      '大文字と小文字を区別する',
      '完全に同一なセルだけを検索する',
      '半角と全角を区別する',
    ]);
    expect(q<HTMLSelectElement>(ctx.host, '#fc-find-look-in').value).toBe('formulas');
    expect(q<HTMLElement>(ctx.host, '.fc-find__field--format').hidden).toBe(true);
    const selected = [
      ...ctx.host.querySelectorAll<HTMLSelectElement>('.fc-find__options select'),
    ].map((s) => s.selectedOptions[0]?.textContent);
    expect(selected).toEqual(['シート', '行', '数式']);
  });

  it('exposes the tab and options state the Mac layout positions by', () => {
    // Excel lays the dialog out per state: 638x226 collapsed, 638x310 with
    // オプション open, and the replace tab narrows the popups to 74 pt.
    ctx.handle.open('find');
    const panel = q<HTMLElement>(ctx.host, '.fc-find');
    expect(panel.dataset.fcFindTab).toBe('find');
    expect(panel.classList.contains('fc-find--options')).toBe(false);
    q<HTMLButtonElement>(ctx.host, '.fc-find__options-btn').click();
    expect(panel.classList.contains('fc-find--options')).toBe(true);
    ctx.handle.open('replace');
    expect(panel.dataset.fcFindTab).toBe('replace');
  });

  it('shows the cell count at the foot from the start', () => {
    ctx.handle.open('find');
    expect(q<HTMLElement>(ctx.host, '.fc-find__status').textContent).toBe('0セルが見つかりました');
  });

  it('lays the replace tab footer out as 置換 / すべて置換 then すべて検索 / 前へ / 次へ / 閉じる', () => {
    ctx.handle.open('replace');
    const labels = [...ctx.host.querySelectorAll<HTMLButtonElement>('.fc-find__footer button')]
      .filter((b) => !b.hidden)
      .map((b) => b.textContent);
    expect(labels).toEqual(['置換', 'すべて置換', 'すべて検索', '前へ', '次へ', '閉じる']);
  });
});

describe('find and replace, half-width / full-width', () => {
  const found = (host: HTMLElement, query = '123'): string => {
    const input = q<HTMLInputElement>(host, '#fc-find-find-input');
    input.value = query;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    q<HTMLButtonElement>(host, '.fc-find__btn--find-all').click();
    return q<HTMLElement>(host, '.fc-find__results-summary').textContent ?? '';
  };

  // Desktop app (Range.Find): letters, digits and plain kana fold their width
  // either way; Match byte only tells ｶﾞ from ガ.
  it('keeps 123 equal to １２３ when 半角と全角を区別する is ticked (Mac)', async () => {
    const { host, handle } = await mount('mac', '１２３');
    handle.open('find');
    expect(found(host)).toBe('1セルが見つかりました');
    q<HTMLButtonElement>(host, '.fc-find__options-btn').click();
    const width = q<HTMLInputElement>(host, '#fc-find-width');
    width.checked = true;
    width.dispatchEvent(new Event('change', { bubbles: true }));
    expect(found(host)).toBe('1セルが見つかりました');
  });

  it('tells ｶﾞ from ガ only while 半角と全角を区別する is ticked (Mac)', async () => {
    const { host, handle } = await mount('mac', 'ｶﾞ');
    handle.open('find');
    expect(found(host, 'ガ')).toBe('1セルが見つかりました');
    q<HTMLButtonElement>(host, '.fc-find__options-btn').click();
    const width = q<HTMLInputElement>(host, '#fc-find-width');
    width.checked = true;
    width.dispatchEvent(new Event('change', { bubbles: true }));
    expect(found(host, 'ガ')).toBe('0セルが見つかりました');
  });

  it('folds width on other platforms and offers no width box', async () => {
    const { host, handle } = await mount('default', '１２３');
    handle.open('find');
    expect(host.querySelector('#fc-find-width')).toBeNull();
    expect(found(host)).toBe('1 個のセルが見つかりました');
  });
});
