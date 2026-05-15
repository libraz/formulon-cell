import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

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
    expect(root?.textContent).toContain('Apply');
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
});
