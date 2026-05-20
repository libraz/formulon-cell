import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
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
    expect(store.getState().ui.filterRange).toEqual(range);
    expect(store.getState().ui.filterCriteria).toEqual([{ range, byCol: 0, hiddenValues: ['B'] }]);
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(store.getState().layout.hiddenRows.size).toBe(0);
    expect(store.getState().ui.filterRange).toBeNull();
    expect(store.getState().ui.filterCriteria).toEqual([]);

    history.redo();
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().ui.filterRange).toEqual(range);

    handle.open(range, 0, { x: 10, y: 20, h: 24 });
    document.querySelector<HTMLButtonElement>('.fc-filter-dropdown__clear')?.click();
    expect(store.getState().layout.hiddenRows.size).toBe(0);
    expect(store.getState().ui.filterRange).toBeNull();

    history.undo();
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(store.getState().ui.filterRange).toEqual(range);
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

  it('keeps action button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/filter-dropdown.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('const createFilterDropdownActionButton');
    expect(source).toContain(
      "const apply = createFilterDropdownActionButton('fc-filter-dropdown__apply', t.apply)",
    );
    expect(source).toContain(
      "const clear = createFilterDropdownActionButton('fc-filter-dropdown__clear', t.clear)",
    );
    expect(source).not.toContain("document.createElement('button')");
  });
});
