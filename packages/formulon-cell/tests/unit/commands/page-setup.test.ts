import { describe, expect, it } from 'vitest';
import {
  clearPrintTitles,
  listPageSetups,
  pageSetupForSheet,
  resetPageSetup,
  setPageSetup,
  setPrintTitleCols,
  setPrintTitleRows,
} from '../../../src/commands/page-setup.js';
import { createSpreadsheetStore, defaultPageSetup } from '../../../src/store/store.js';

describe('page setup commands', () => {
  it('reads defaults and lists only explicitly configured sheets', () => {
    const store = createSpreadsheetStore();

    expect(pageSetupForSheet(store.getState(), 0)).toEqual(defaultPageSetup());
    expect(listPageSetups(store.getState())).toEqual([]);

    const setup = setPageSetup(store, 2, {
      orientation: 'landscape',
      margins: { top: 1 },
    });

    expect(setup.orientation).toBe('landscape');
    expect(setup.margins).toEqual({ ...defaultPageSetup().margins, top: 1 });
    expect(listPageSetups(store.getState())).toEqual([{ sheet: 2, setup }]);
  });

  it('resets a sheet back to defaults', () => {
    const store = createSpreadsheetStore();
    setPageSetup(store, 0, { paperSize: 'letter' });

    expect(resetPageSetup(store, 0)).toEqual(defaultPageSetup());
    expect(listPageSetups(store.getState())).toEqual([]);
  });

  it('sets, validates, and clears print titles', () => {
    const store = createSpreadsheetStore();

    expect(setPrintTitleRows(store, 0, ' 1:3 ')).toMatchObject({ printTitleRows: '1:3' });
    expect(setPrintTitleCols(store, 0, '$A:$B')).toMatchObject({ printTitleCols: '$A:$B' });
    expect(setPrintTitleRows(store, 0, 'abc')).toBeNull();
    expect(setPrintTitleCols(store, 0, '1:3')).toBeNull();

    const cleared = clearPrintTitles(store, 0);
    expect(cleared.printTitleRows).toBeUndefined();
    expect(cleared.printTitleCols).toBeUndefined();
  });
});
