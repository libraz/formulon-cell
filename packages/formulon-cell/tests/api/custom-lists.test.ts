import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { workbookHandleOf } from '../../src/api/workbook.js';
import { customListsSourceFor } from '../../src/commands/custom-lists.js';
import { fillRange } from '../../src/commands/fill.js';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { makeHost } from './helpers.js';

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

const mountSheet = async (options: Parameters<typeof mount>[1] = {}) => {
  const instance = await mount(makeHost(), { locale: 'en', ...options });
  live.push(instance);
  return instance;
};

describe('customLists option', () => {
  it('holds the user lists, reports a user change and leaves the built-in lists out', async () => {
    const sheet = await mountSheet({ customLists: [['Low', 'Mid', 'High']] });
    const updates = vi.fn();
    sheet.on('update', updates);
    const source = customListsSourceFor(spreadsheetInternals(sheet).store);
    expect(source?.get()).toEqual([['Low', 'Mid', 'High']]);

    source?.set([
      ['Low', 'Mid', 'High'],
      ['North', 'South'],
    ]);
    expect(sheet.options.customLists).toEqual([
      ['Low', 'Mid', 'High'],
      ['North', 'South'],
    ]);
    expect(updates).toHaveBeenLastCalledWith({ keys: ['customLists'], source: 'user' });

    sheet.update({ customLists: null });
    expect(source?.get()).toEqual([]);
    expect(updates).toHaveBeenLastCalledWith(
      expect.objectContaining({ keys: ['customLists'], source: 'api' }),
    );
  });

  it('lets AutoFill continue a user list', async () => {
    const sheet = await mountSheet({ customLists: [['Low', 'Mid', 'High']] });
    const { store } = spreadsheetInternals(sheet);
    sheet.workbook.sheet(0).setText(0, 0, 'Low');
    sheet.workbook.sheet(0).setText(1, 0, 'Mid');
    fillRange(
      store.getState(),
      workbookHandleOf(sheet.workbook),
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      { store },
    );
    expect(sheet.workbook.sheet(0).getValue(2, 0)).toBe('High');
  });
});
