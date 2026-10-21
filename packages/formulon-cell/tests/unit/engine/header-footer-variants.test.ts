import { describe, expect, it } from 'vitest';
import {
  hydratePageSetupFromEngine,
  syncPageSetupToEngine,
} from '../../../src/engine/print-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, getPageSetup, mutators } from '../../../src/store/store.js';

describe('first-page and even-page header / footer persistence', () => {
  it('saves and reloads the first and even sections with their flags', async () => {
    const wb = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      expect(wb.capabilities.printSettings).toBe(true);
      const store = createSpreadsheetStore();
      mutators.setPageSetup(store, 0, {
        headerCenter: 'Odd title',
        differentFirstPage: true,
        differentOddEvenPages: true,
        firstHeaderCenter: 'Cover & title',
        firstFooterRight: 'Page &P',
        evenHeaderLeft: 'Even left',
        evenFooterCenter: 'Even &N',
      });
      syncPageSetupToEngine(wb, store, 0);

      loaded = await WorkbookHandle.loadBytes(wb.save());
      expect(loaded.getSheetHeaderFooter(0)).toMatchObject({
        oddHeader: '&COdd title',
        firstHeader: '&CCover & title',
        firstFooter: '&RPage &P',
        evenHeader: '&LEven left',
        evenFooter: '&CEven &N',
        differentFirst: true,
        differentOddEven: true,
      });

      const fresh = createSpreadsheetStore();
      hydratePageSetupFromEngine(loaded, fresh, 0);
      expect(getPageSetup(fresh.getState(), 0)).toMatchObject({
        headerCenter: 'Odd title',
        firstHeaderCenter: 'Cover & title',
        firstFooterRight: 'Page &P',
        evenHeaderLeft: 'Even left',
        evenFooterCenter: 'Even &N',
        differentFirstPage: true,
        differentOddEvenPages: true,
      });
    } finally {
      loaded?.dispose();
      wb.dispose();
    }
  });
});
