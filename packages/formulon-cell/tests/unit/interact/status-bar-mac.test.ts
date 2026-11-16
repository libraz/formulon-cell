import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachStatusBar } from '../../../src/interact/status-bar.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const seed = (platform: 'mac' | null) => {
  const host = document.createElement('div');
  host.className = 'fc-host';
  if (platform) host.dataset.fcPlatform = platform;
  const bar = document.createElement('div');
  bar.className = 'fc-host__statusbar';
  host.appendChild(bar);
  document.body.appendChild(host);
  const store = createSpreadsheetStore();
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
      value: { kind: 'number', value: 1 },
      formula: null,
    });
    cells.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
      value: { kind: 'number', value: 2 },
      formula: null,
    });
    return {
      ...s,
      data: { ...s.data, cells },
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      },
    };
  });
  const handle = attachStatusBar({
    getLocale: () => 'en-US',
    statusbar: bar,
    store,
    getEngineLabel: () => 'engine',
    strings: ja,
  });
  handle.refresh?.();
  return { bar, handle };
};

describe('status bar aggregates layout', () => {
  beforeEach(() => {
    document.body.replaceChildren();
  });
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('keeps each Mac aggregate as its own item without a separator or cell count', () => {
    const { bar, handle } = seed('mac');
    const items = Array.from(
      bar.querySelectorAll('.fc-host__statusbar-agg'),
      (el) => el.textContent,
    );
    expect(items).toEqual(['平均: 1.5', 'データの個数: 2', '合計: 3']);
    expect(bar.querySelector('.fc-host__statusbar-aggs')?.textContent).not.toContain('·');
    expect(bar.querySelector('.fc-host__statusbar-right')?.textContent).toBe('');
    handle.detach?.();
  });

  it('joins the other profiles aggregates with a dot and shows the cell count', () => {
    const { bar, handle } = seed(null);
    expect(bar.querySelector('.fc-host__statusbar-aggs')?.textContent).toBe(
      '平均: 1.5 · データの個数: 2 · 合計: 3',
    );
    expect(bar.querySelector('.fc-host__statusbar-right')?.textContent).toContain('セル');
    handle.detach?.();
  });
});

describe('status bar view buttons on the Mac host', () => {
  beforeEach(() => document.body.replaceChildren());

  it('names the page break button 改ページのプレビュー', () => {
    const view = (platform: 'mac' | null): HTMLElement | null => {
      document.body.replaceChildren();
      seed(platform);
      return document.querySelector<HTMLElement>('[aria-label^="改ページ"]');
    };
    const mac = view('mac');
    expect(mac?.getAttribute('aria-label')).toBe('改ページのプレビュー');
    expect(mac?.title).toBe('改ページのプレビュー');
    expect(view(null)?.getAttribute('aria-label')).toBe(ja.statusBar.pageBreakPreview);
  });
});

describe('status bar clicks on the Mac host', () => {
  beforeEach(() => document.body.replaceChildren());

  it('copies an aggregate value on a left click without opening the menu', async () => {
    const writeText = vi.fn(async () => undefined);
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true });
    const { bar } = seed('mac');
    const sum = Array.from(bar.querySelectorAll<HTMLElement>('.fc-host__statusbar-agg')).find(
      (el) => el.textContent?.startsWith('合計'),
    );
    sum?.click();
    expect(writeText).toHaveBeenCalledWith('3');
    expect(document.querySelector<HTMLElement>('.fc-statusbar__chooser')?.style.display).toBe(
      'none',
    );
  });

  it('opens the Zoom dialog from the percent label', async () => {
    const { bar } = seed('mac');
    bar.querySelector<HTMLElement>('.fc-host__statusbar-zoom-label')?.click();
    await Promise.resolve();
    expect(document.body.textContent).toContain('選択範囲をズーム');
  });

  it('lists the Mac status bar customisation menu in the desktop order', () => {
    const { bar } = seed('mac');
    bar.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    const rows = Array.from(chooser?.children ?? [], (el) =>
      el.getAttribute('role') === 'separator' ? '-' : el.textContent?.trim(),
    );
    expect(rows).toEqual([
      'ステータス バーのユーザー設定',
      'CapsLock',
      'マクロの記録',
      '-',
      '平均',
      'データの個数',
      '数値の個数',
      '最小値',
      '最大値',
      '合計',
      '-',
      'アップロード状態',
      '-',
      '表示選択ショートカット',
      'ズーム スライダー',
      'ズーム',
    ]);
  });
});

describe('status bar filter record count', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('shows how many records a filter found and hides the count when cleared', () => {
    const { bar, handle } = seed('mac');
    const filterText = () => bar.querySelector<HTMLElement>('.fc-host__statusbar-filter');
    expect(filterText()?.hidden).toBe(true);

    const range = { sheet: 0, r0: 0, c0: 0, r1: 9, c1: 1 };
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        filterRanges: [range],
        filterCriteria: [{ range, byCol: 1, hiddenValues: ['Y'] }],
      },
      layout: { ...s.layout, hiddenRows: new Set([7, 8, 9]) },
    }));
    const filteredBar = document.createElement('div');
    document.body.appendChild(filteredBar);
    const filteredHandle = attachStatusBar({
      getLocale: () => 'en-US',
      statusbar: filteredBar,
      store,
      getEngineLabel: () => 'engine',
      strings: ja,
    });
    filteredHandle.refresh?.();
    const text = filteredBar.querySelector<HTMLElement>('.fc-host__statusbar-filter');
    expect(text?.hidden).toBe(false);
    expect(text?.textContent).toBe('9 レコード中 6 個が見つかりました');

    store.setState((s) => ({ ...s, ui: { ...s.ui, filterCriteria: [] } }));
    filteredHandle.refresh?.();
    expect(text?.hidden).toBe(true);
    handle.detach?.();
  });
});
