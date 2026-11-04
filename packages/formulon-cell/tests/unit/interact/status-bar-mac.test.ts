import { afterEach, beforeEach, describe, expect, it } from 'vitest';
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
