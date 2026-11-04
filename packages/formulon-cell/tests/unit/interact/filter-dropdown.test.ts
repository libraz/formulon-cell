import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { applyValueFilter } from '../../../src/commands/filter.js';
import { History } from '../../../src/commands/history.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { ensureOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('attachFilterDropdown', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders localized default labels', () => {
    const handle = attachFilterDropdown({ store });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, 0, { x: 10, y: 20, h: 24 });

    const root = document.querySelector<HTMLElement>('.fc-filter-dropdown');
    expect(root?.getAttribute('aria-label')).toBe('フィルター');
    expect(root?.getAttribute('aria-modal')).toBe('false');
    const search = root?.querySelector<HTMLInputElement>('.fc-filter-dropdown__search');
    expect(search?.placeholder).toBe('検索…');
    expect(search?.getAttribute('aria-label')).toBe('検索…');
    expect(
      root?.querySelector<HTMLElement>('.fc-filter-dropdown__list')?.getAttribute('role'),
    ).toBe('group');
    expect(
      root?.querySelector<HTMLElement>('.fc-filter-dropdown__list')?.getAttribute('aria-label'),
    ).toBe('フィルター');
    expect(root?.textContent).toContain('(すべて選択)');
    expect(root?.textContent).toContain('OK');
    expect(root?.textContent).toContain('クリア');
    const condition = root?.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
    expect(condition?.getAttribute('aria-label')).toBe('条件でフィルター');
    expect(Array.from(condition?.options ?? []).map((option) => option.value)).toEqual([
      '',
      'equals',
      'notEquals',
      'contains',
      'notContains',
      'greaterThan',
      'greaterThanOrEqual',
      'lessThan',
      'lessThanOrEqual',
    ]);
    handle.detach();
  });

  it('keeps focus in the condition input when it is focused before the search focus frame', async () => {
    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, 0, { x: 10, y: 20, h: 24 });
    const condition = document.querySelector<HTMLInputElement>(
      '.fc-filter-dropdown__condition-value',
    );
    condition?.focus();
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    expect(document.activeElement).toBe(condition);
    handle.detach();
  });

  it('accepts an English dictionary override', () => {
    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, 0, { x: 10, y: 20, h: 24 });

    const root = document.querySelector<HTMLElement>('.fc-filter-dropdown');
    expect(root?.getAttribute('aria-label')).toBe('Filter');
    expect(root?.querySelector<HTMLInputElement>('.fc-filter-dropdown__search')?.placeholder).toBe(
      'Search…',
    );
    expect(root?.textContent).toContain('(Select all)');
    expect(root?.textContent).toContain('Filter by condition');
    expect(root?.textContent).toContain('Apply');
    handle.detach();
  });

  it('uses the supplied host portal when multiple spreadsheets share a document', () => {
    const firstHost = document.createElement('div');
    firstHost.className = 'fc-host';
    const secondHost = document.createElement('div');
    secondHost.className = 'fc-host';
    const firstRoot = document.createElement('section');
    const secondRoot = document.createElement('section');
    document.body.append(firstHost, secondHost, firstRoot, secondRoot);
    ensureOverlayPortal(firstHost, { root: firstRoot });
    ensureOverlayPortal(secondHost, { root: secondRoot });

    const handle = attachFilterDropdown({ store, strings: en, host: secondHost });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, 0, { x: 10, y: 20, h: 24 });

    expect(secondRoot.querySelector('.fc-filter-dropdown')).not.toBeNull();
    expect(firstRoot.querySelector('.fc-filter-dropdown')).toBeNull();
    handle.detach();
  });

  it('applies a condition filter when a condition and value are entered', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        value: { kind: 'text', value: 'paper' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 2, col: 0 }), {
        value: { kind: 'text', value: 'ink' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 3, col: 0 }), {
        value: { kind: 'text', value: 'pencil' },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, 0, { x: 10, y: 20, h: 24 });

    const op = document.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
    const value = document.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value');
    if (op) op.value = 'contains';
    if (value) value.value = 'p';
    document.querySelector<HTMLButtonElement>('.fc-filter-dropdown__apply')?.click();

    expect(store.getState().layout.hiddenRows.has(1)).toBe(false);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(false);
    handle.detach();
  });

  it('applies with Enter while focus is inside the dropdown', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        value: { kind: 'text', value: 'A' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 2, col: 0 }), {
        value: { kind: 'text', value: 'B' },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }, 0, { x: 10, y: 20, h: 24 });
    const checkboxes = document.querySelectorAll<HTMLInputElement>(
      '.fc-filter-dropdown__row:not(.fc-filter-dropdown__row--all) input[type="checkbox"]',
    );
    const second = checkboxes[1];
    if (!second) throw new Error('expected second filter checkbox');
    second.checked = false;
    second.dispatchEvent(new Event('change', { bubbles: true }));
    second.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(handle.isOpen()).toBe(false);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
  });

  it('records apply and clear through unified history when provided', () => {
    const history = new History();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        value: { kind: 'text', value: 'A' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 2, col: 0 }), {
        value: { kind: 'text', value: 'B' },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    const handle = attachFilterDropdown({ store, strings: en, history });
    handle.open(range, 0, { x: 10, y: 20, h: 24 });
    const second = document.querySelectorAll<HTMLInputElement>(
      '.fc-filter-dropdown__row:not(.fc-filter-dropdown__row--all) input[type="checkbox"]',
    )[1];
    if (!second) throw new Error('expected second filter checkbox');
    second.checked = false;
    second.dispatchEvent(new Event('change', { bubbles: true }));
    document.querySelector<HTMLButtonElement>('.fc-filter-dropdown__apply')?.click();

    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().ui.filterRanges).toEqual([range]);
    expect(store.getState().ui.filterCriteria).toEqual([{ range, byCol: 0, hiddenValues: ['B'] }]);
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(store.getState().layout.hiddenRows.size).toBe(0);
    expect(store.getState().ui.filterRanges).toEqual([]);
    expect(store.getState().ui.filterCriteria).toEqual([]);

    history.redo();
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().ui.filterRanges).toEqual([range]);

    handle.open(range, 0, { x: 10, y: 20, h: 24 });
    document.querySelector<HTMLButtonElement>('.fc-filter-dropdown__clear')?.click();
    // Clear drops the criteria; the filter buttons stay, as in the desktop app.
    expect(store.getState().layout.hiddenRows.size).toBe(0);
    expect(store.getState().ui.filterRanges).toEqual([range]);
    expect(store.getState().ui.filterCriteria).toEqual([]);

    history.undo();
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().ui.filterRanges).toEqual([range]);
    handle.detach();
  });

  it('restores checklist state from column criteria instead of hidden rows from other columns', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        value: { kind: 'text', value: 'Region' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 0, col: 1 }), {
        value: { kind: 'text', value: 'Type' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        value: { kind: 'text', value: 'West' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 1, col: 1 }), {
        value: { kind: 'text', value: 'A' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 2, col: 0 }), {
        value: { kind: 'text', value: 'East' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 2, col: 1 }), {
        value: { kind: 'text', value: 'B' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 3, col: 0 }), {
        value: { kind: 'text', value: 'East' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 3, col: 1 }), {
        value: { kind: 'text', value: 'A' },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
    applyValueFilter(store.getState(), store, range, 1, ['B']);

    const handle = attachFilterDropdown({ store, strings: en });
    handle.open(range, 0, { x: 10, y: 20, h: 24 });
    document.querySelector<HTMLButtonElement>('.fc-filter-dropdown__apply')?.click();

    expect(store.getState().ui.filterCriteria).toEqual([{ range, byCol: 1, hiddenValues: ['B'] }]);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(false);
    handle.detach();
  });

  it('restores active condition criteria into the condition controls', () => {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        filterCriteria: [
          {
            range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
            byCol: 0,
            hiddenValues: [],
            condition: { op: 'contains', value: 'pen' },
          },
        ],
      },
    }));
    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, 0, { x: 10, y: 20, h: 24 });

    expect(
      document.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op')?.value,
    ).toBe('contains');
    expect(
      document.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value')?.value,
    ).toBe('pen');
    handle.detach();
  });

  it('formats canonical numeric criteria in the active locale', () => {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        filterCriteria: [
          {
            range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 },
            byCol: 0,
            hiddenValues: [],
            condition: { op: 'greaterThan', value: '1.5' },
          },
        ],
      },
    }));
    const handle = attachFilterDropdown({ store, strings: en, locale: 'de-DE' });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, 0, { x: 10, y: 20, h: 24 });

    expect(
      document.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value')?.value,
    ).toBe('1,5');
    handle.detach();
  });

  it('restores focus to the opener when dismissed with Escape', () => {
    const opener = document.createElement('button');
    document.body.appendChild(opener);
    opener.focus();

    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, 0, { x: 10, y: 20, h: 24 });
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(handle.isOpen()).toBe(false);
    expect(document.activeElement).toBe(opener);
  });

  it('supports Excel-style arrow navigation from search into checkbox rows', () => {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        value: { kind: 'text', value: 'A' },
        formula: null,
      });
      cells.set(addrKey({ sheet: 0, row: 2, col: 0 }), {
        value: { kind: 'text', value: 'B' },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });

    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }, 0, { x: 10, y: 20, h: 24 });

    const search = document.querySelector<HTMLInputElement>('.fc-filter-dropdown__search');
    const boxes = Array.from(
      document.querySelectorAll<HTMLInputElement>(
        '.fc-filter-dropdown__list input[type="checkbox"]',
      ),
    );
    expect(boxes).toHaveLength(3);
    search?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(boxes[0]);

    boxes[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(boxes[1]);

    boxes[1]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(document.activeElement).toBe(boxes[2]);

    boxes[2]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(boxes[0]);

    handle.detach();
  });

  it('keeps the dropdown within the viewport near the bottom-right edge', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 240 });

    const handle = attachFilterDropdown({ store, strings: en });
    handle.open({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }, 0, { x: 310, y: 220, h: 20 });

    const root = document.querySelector<HTMLElement>('.fc-filter-dropdown');
    expect(root).not.toBeNull();
    if (!root) throw new Error('missing filter dropdown');

    expect(root.style.left).toBe('56px');
    expect(root.style.top).toBe('4px');
    handle.detach();
  });

  it('keeps the filter dropdown close to Excel 365 desktop AutoFilter geometry', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/popups/filter-dropdown.css'), 'utf8');

    expect(css).toMatch(
      /\.fc-filter-dropdown\s*\{[\s\S]*?border-radius: 2px;[\s\S]*?box-shadow:[\s\S]*?font-size: 12px;[\s\S]*?width: 240px;/,
    );
    expect(css).toMatch(
      /\.fc-filter-dropdown__search\s*\{[\s\S]*?margin: 7px 8px 6px;[\s\S]*?padding: 4px 7px;[\s\S]*?border-radius: 2px;/,
    );
    expect(css).toMatch(
      /\.fc-filter-dropdown__search:focus\s*\{[\s\S]*?box-shadow: inset 0 0 0 1px var\(--fc-accent/,
    );
    expect(css).toMatch(
      /\.fc-filter-dropdown__condition-op,[\s\S]*?\.fc-filter-dropdown__condition-value\s*\{[\s\S]*?padding: 4px 7px;[\s\S]*?border-radius: 2px;/,
    );
    expect(css).toMatch(
      /\.fc-filter-dropdown__row\s*\{[\s\S]*?gap: 6px;[\s\S]*?min-height: 24px;[\s\S]*?padding: 2px 8px;/,
    );
    expect(css).toMatch(
      /\.fc-filter-dropdown__row:hover\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(css).toMatch(
      /\.fc-filter-dropdown__actions button\s*\{[\s\S]*?padding: 3px 12px;[\s\S]*?border-radius: 2px;/,
    );
    expect(css).not.toContain('background: var(--fc-accent-soft, rgba(30, 122, 224, 0.1))');
  });

  it('keeps action button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/filter-dropdown.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('const createFilterDropdownActionButton');
    expect(source).toContain(
      "const apply = createFilterDropdownActionButton(\n      'fc-filter-dropdown__apply'",
    );
    expect(source).toContain(
      "const clear = createFilterDropdownActionButton(\n      'fc-filter-dropdown__clear'",
    );
    expect(source).not.toContain("document.createElement('button')");
  });

  describe('on the Mac host', () => {
    const openMac = () => {
      const host = document.createElement('div');
      host.className = 'fc-host';
      host.dataset.fcPlatform = 'mac';
      document.body.appendChild(host);
      store.setState((state) => {
        const cells = new Map(state.data.cells);
        const put = (
          row: number,
          value: { kind: 'text'; value: string } | { kind: 'number'; value: number },
        ) => cells.set(addrKey({ sheet: 0, row, col: 0 }), { value, formula: null });
        put(0, { kind: 'text', value: 'N' });
        put(1, { kind: 'number', value: 1 });
        put(2, { kind: 'number', value: 3 });
        put(3, { kind: 'number', value: 5 });
        return { ...state, data: { ...state.data, cells } };
      });
      const handle = attachFilterDropdown({ store, host, strings: ja });
      handle.open({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, 0, { x: 10, y: 20, h: 24 });
      return handle;
    };

    it('titles the panel with the column header and lists the numeric conditions', () => {
      const handle = openMac();
      const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown');
      expect(panel?.querySelector('.fc-filter-dropdown__title')?.textContent).toBe('N');
      const options = Array.from(
        panel?.querySelectorAll<HTMLOptionElement>('.fc-filter-dropdown__condition-op option') ??
          [],
        (o) =>
          `${o.getAttribute('aria-disabled') === 'true' || o.disabled ? '[DIS]' : ''}${o.textContent}`,
      );
      expect(options).toEqual([
        '1 つ選択してください',
        '指定の値に等しい',
        '指定の値に等しくない',
        '指定の値より大きい',
        '指定の値以上',
        '指定の値より小さい',
        '指定の値以下',
        '指定の範囲内',
        '上位 10 項目',
        '下位 10 項目',
        '平均より上',
        '平均より下',
      ]);
      handle.detach();
    });

    it('stacks the sort and filter sections and offers colour choices in both', () => {
      const handle = openMac();
      const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown');
      const order = Array.from(panel?.children ?? [], (el) => el.className.split(' ')[0]);
      expect(order.slice(0, 9)).toEqual([
        'fc-filter-dropdown__close',
        'fc-filter-dropdown__title',
        'fc-filter-dropdown__section',
        'fc-filter-dropdown__sort',
        'fc-filter-dropdown__color',
        'fc-filter-dropdown__rule',
        'fc-filter-dropdown__section',
        'fc-filter-dropdown__color',
        'fc-filter-dropdown__condition',
      ]);
      expect(
        Array.from(
          panel?.querySelectorAll('.fc-filter-dropdown__section') ?? [],
          (e) => e.textContent,
        ),
      ).toEqual(['並べ替え', 'フィルター']);
      for (const select of panel?.querySelectorAll<HTMLSelectElement>(
        '.fc-filter-dropdown__color select',
      ) ?? []) {
        expect(Array.from(select.options, (o) => o.textContent)).toEqual([
          'なし',
          'セルの色',
          'フォントの色',
          'セルのアイコン',
        ]);
      }
      handle.detach();
    });

    it('stays open on an outside click and closes with its close button', () => {
      const handle = openMac();
      document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      expect(handle.isOpen()).toBe(true);
      document
        .querySelector<HTMLButtonElement>('.fc-filter-dropdown__close')
        ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      expect(handle.isOpen()).toBe(false);
    });

    it('lists the text conditions without the number-only entries', () => {
      const host = document.createElement('div');
      host.className = 'fc-host';
      host.dataset.fcPlatform = 'mac';
      document.body.appendChild(host);
      store.setState((state) => {
        const cells = new Map(state.data.cells);
        for (const [row, v] of ['h', 'a', 'b'].entries()) {
          cells.set(addrKey({ sheet: 0, row, col: 0 }), {
            value: { kind: 'text', value: v },
            formula: null,
          });
        }
        return { ...state, data: { ...state.data, cells } };
      });
      const handle = attachFilterDropdown({ store, host, strings: ja });
      handle.open({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 }, 0, { x: 10, y: 20, h: 24 });
      const options = Array.from(
        document.querySelectorAll<HTMLOptionElement>('.fc-filter-dropdown__condition-op option'),
        (o) => o.textContent,
      );
      expect(options).toEqual([
        '1 つ選択してください',
        '指定の値に等しい',
        '指定の値に等しくない',
        'を含む',
        'を含まない',
      ]);
      handle.detach();
    });

    it('sorts the column from the panel and greys the colour popups', async () => {
      const wb = await WorkbookHandle.createDefault();
      try {
        wb.setText({ sheet: 0, row: 0, col: 0 }, 'N');
        [3, 1, 2].forEach((v, k) => {
          wb.setNumber({ sheet: 0, row: k + 1, col: 0 }, v);
        });
        const host = document.createElement('div');
        host.className = 'fc-host';
        host.dataset.fcPlatform = 'mac';
        document.body.appendChild(host);
        store.setState((state) => ({
          ...state,
          data: { ...state.data, cells: new Map() },
        }));
        const { mutators } = await import('../../../src/store/store.js');
        mutators.replaceCells(store, wb.cells(0));
        const handle = attachFilterDropdown({
          store,
          host,
          strings: ja,
          getWb: () => wb,
          history: new History(),
        });
        handle.open({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, 0, { x: 10, y: 20, h: 24 });
        const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown');
        expect(panel?.querySelector('.fc-filter-dropdown__sort-asc')?.textContent).toBe('昇順');
        const colors = panel?.querySelectorAll<HTMLSelectElement>(
          '.fc-filter-dropdown__color select',
        );
        expect(colors?.length).toBe(2);
        expect(Array.from(colors ?? []).every((c) => c.disabled)).toBe(true);
        panel?.querySelector<HTMLButtonElement>('.fc-filter-dropdown__sort-asc')?.click();
        expect([1, 2, 3].map((r) => wb.getValue({ sheet: 0, row: r, col: 0 }))).toEqual([
          { kind: 'number', value: 1 },
          { kind: 'number', value: 2 },
          { kind: 'number', value: 3 },
        ]);
        expect(handle.isOpen()).toBe(true);
      } finally {
        wb.dispose();
      }
    });

    it('applies a checklist toggle at once and stays open after clearing', () => {
      const handle = openMac();
      const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown');
      const boxes = () =>
        Array.from(
          document.querySelectorAll<HTMLInputElement>(
            '.fc-filter-dropdown__list input[type="checkbox"]',
          ),
        );
      const second = boxes()[2];
      if (second) {
        second.checked = false;
        second.dispatchEvent(new Event('change', { bubbles: true }));
      }
      expect(store.getState().ui.filterCriteria).toHaveLength(1);
      expect([...store.getState().layout.hiddenRows]).toEqual([2]);

      panel?.querySelector<HTMLButtonElement>('.fc-filter-dropdown__clear')?.click();
      expect(handle.isOpen()).toBe(true);
      expect(store.getState().ui.filterCriteria).toHaveLength(0);
      expect(store.getState().layout.hiddenRows.size).toBe(0);
      expect(boxes().every((b) => b.checked)).toBe(true);
      handle.detach();
    });

    it('filters the top 10 with an inline count and unit', () => {
      const handle = openMac();
      const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown');
      const op = panel?.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
      const count = panel?.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value');
      const unit = panel?.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-unit');
      const upper = panel?.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value2');
      expect(unit?.hidden).toBe(true);
      if (op) {
        op.value = 'top10';
        op.dispatchEvent(new Event('change'));
      }
      expect(count?.hidden).toBe(false);
      expect(count?.value).toBe('10');
      expect(unit?.hidden).toBe(false);
      expect(upper?.hidden).toBe(true);
      expect(Array.from(unit?.options ?? [], (o) => o.textContent)).toEqual(['項目', 'パーセント']);
      if (count) {
        count.value = '1';
        count.dispatchEvent(new Event('change'));
      }
      expect([...store.getState().layout.hiddenRows].sort()).toEqual([1, 2]);
      handle.detach();
    });

    it('asks for two bounds on a range and none on an average', () => {
      const handle = openMac();
      const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown');
      const op = panel?.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
      const low = panel?.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value');
      const high = panel?.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value2');
      if (op) {
        op.value = 'between';
        op.dispatchEvent(new Event('change'));
      }
      expect([low?.hidden, high?.hidden]).toEqual([false, false]);
      if (low && high) {
        low.value = '2';
        low.dispatchEvent(new Event('change'));
        expect(store.getState().ui.filterCriteria).toEqual([]);
        high.value = '4';
        high.dispatchEvent(new Event('change'));
      }
      expect([...store.getState().layout.hiddenRows].sort()).toEqual([1, 3]);
      if (op) {
        op.value = 'aboveAverage';
        op.dispatchEvent(new Event('change'));
      }
      expect([low?.hidden, high?.hidden]).toEqual([true, true]);
      expect([...store.getState().layout.hiddenRows].sort()).toEqual([1, 2]);
      handle.detach();
    });

    it('auto-applies a condition while the panel stays open, and Apply waits for auto-apply off', () => {
      const handle = openMac();
      const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown');
      const op = panel?.querySelector<HTMLSelectElement>('.fc-filter-dropdown__condition-op');
      const value = panel?.querySelector<HTMLInputElement>('.fc-filter-dropdown__condition-value');
      const apply = panel?.querySelector<HTMLButtonElement>('.fc-filter-dropdown__apply');
      const auto = panel?.querySelector<HTMLInputElement>('.fc-filter-dropdown__auto input');
      expect(value?.hidden).toBe(true);
      expect(auto?.checked).toBe(true);
      expect(apply?.disabled).toBe(true);
      if (op) {
        op.value = 'greaterThan';
        op.dispatchEvent(new Event('change'));
      }
      expect(value?.hidden).toBe(false);
      if (value) {
        value.value = '2';
        value.dispatchEvent(new Event('change'));
      }
      expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
      expect(store.getState().layout.hiddenRows.has(2)).toBe(false);
      expect(document.querySelector('.fc-filter-dropdown')).not.toBeNull();
      if (auto) {
        auto.checked = false;
        auto.dispatchEvent(new Event('change'));
      }
      expect(apply?.disabled).toBe(false);
      handle.detach();
    });
  });
});
