import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

describe('toolbar contextual tab sync', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'ja' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('skips the pivot lookup on store changes that leave its inputs untouched', () => {
    const store = sheet.instance.store;
    const lookup = vi.spyOn(sheet.workbook, 'getPivotTables');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });
    lookup.mockClear();

    store.setState((s) => ({ ...s, ui: { ...s.ui, hover: { sheet: 0, row: 1, col: 1 } } }));
    store.setState((s) => ({ ...s, ui: { ...s.ui, hover: null } }));
    // The ribbon state projection does its own lookup per notification; the
    // contextual-tab sync must add none of its own.
    expect(lookup.mock.calls.length).toBeLessThanOrEqual(2);

    mutators.setActive(store, { sheet: 0, row: 3, col: 2 });
    expect(lookup.mock.calls.length).toBeGreaterThan(2);
    tb.dispose();
  });
});
