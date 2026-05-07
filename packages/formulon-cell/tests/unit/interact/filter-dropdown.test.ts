import { afterEach, beforeEach, describe, expect, it } from 'vitest';
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
    expect(root?.querySelector<HTMLInputElement>('.fc-filter-dropdown__search')?.placeholder).toBe(
      '検索…',
    );
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
});
