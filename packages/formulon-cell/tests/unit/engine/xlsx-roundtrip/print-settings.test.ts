import { describe, expect, it } from 'vitest';
import { clearPrintArea, setPrintArea } from '../../../../src/commands/page-setup.js';
import {
  hydratePageSetupFromEngine,
  syncPageSetupToEngine,
} from '../../../../src/engine/print-sync.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, getPageSetup } from '../../../../src/store/store.js';

import { canLoadWasm } from './fixtures.js';

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
  it('saves and reloads worksheet print settings when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.printSettings) return;
      const store = createSpreadsheetStore();
      store.setState((state) => ({
        ...state,
        pageSetup: {
          setupBySheet: new Map([
            [
              0,
              {
                ...getPageSetup(state, 0),
                orientation: 'landscape' as const,
                paperSize: 'letter' as const,
                scale: 0.8,
                margins: { top: 1, right: 0.4, bottom: 1, left: 0.4 },
                headerMargin: 0.2,
                footerMargin: 0.2,
                showGridlines: true,
                centerHorizontally: true,
                headerCenter: 'Quarterly report',
                footerRight: 'Page &P of &N',
                printArea: 'A1:D10',
                printTitleRows: '1:1',
                printTitleCols: 'A:A',
                manualPageBreakRows: [12],
                manualPageBreakCols: [3],
              },
            ],
          ]),
        },
      }));
      syncPageSetupToEngine(first, store, 0);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const restored = createSpreadsheetStore();
        hydratePageSetupFromEngine(reloaded, restored, 0);
        const setup = getPageSetup(restored.getState(), 0);

        expect(setup.orientation).toBe('landscape');
        expect(setup.paperSize).toBe('letter');
        expect(setup.scale).toBeCloseTo(0.8);
        expect(setup.margins).toEqual({ top: 1, right: 0.4, bottom: 1, left: 0.4 });
        expect(setup.headerMargin).toBeCloseTo(0.2);
        expect(setup.showGridlines).toBe(true);
        expect(setup.centerHorizontally).toBe(true);
        expect(setup.headerCenter).toBe('Quarterly report');
        expect(setup.footerRight).toBe('Page &P of &N');
        expect(setup.printArea).toBe('A1:D10');
        expect(setup.printTitleRows).toBe('1:1');
        expect(setup.printTitleCols).toBe('A:A');
        if (reloaded.capabilities.pageBreaks) {
          expect(setup.manualPageBreakRows).toEqual([12]);
          expect(setup.manualPageBreakCols).toEqual([3]);
        }
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('clears a saved print area when the user removes it', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.printSettings) return;
      const store = createSpreadsheetStore();
      setPrintArea(store, 0, 'A1:C5');
      syncPageSetupToEngine(first, store, 0);
      expect(first.getSheetPrintArea(0)).toBe('A1:C5');

      clearPrintArea(store, 0);
      syncPageSetupToEngine(first, store, 0);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        expect(reloaded.getSheetPrintArea(0) ?? '').toBe('');
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads the print attributes the raw fragment carries', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.printSettingsXml) return;
      // An attribute this package does not model at all, seeded through the
      // fragment so the merge can be shown to leave it alone.
      expect(first.setSheetPageSetupXml(0, '<pageSetup copies="3"/>')).toBe(true);

      const store = createSpreadsheetStore();
      store.setState((state) => ({
        ...state,
        pageSetup: {
          setupBySheet: new Map([
            [
              0,
              {
                ...getPageSetup(state, 0),
                blackAndWhite: true,
                draftQuality: true,
                comments: 'endOfSheet' as const,
                cellErrorsAs: 'na' as const,
                pageOrder: 'overThenDown' as const,
                firstPageNumber: 4,
                printQuality: '600' as const,
              },
            ],
          ]),
        },
      }));
      syncPageSetupToEngine(first, store, 0);
      expect(first.getSheetPageSetupXml(0)).toContain('copies="3"');

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const restored = createSpreadsheetStore();
        hydratePageSetupFromEngine(reloaded, restored, 0);
        const setup = getPageSetup(restored.getState(), 0);

        expect(setup.blackAndWhite).toBe(true);
        expect(setup.draftQuality).toBe(true);
        expect(setup.comments).toBe('endOfSheet');
        expect(setup.cellErrorsAs).toBe('na');
        expect(setup.pageOrder).toBe('overThenDown');
        expect(setup.firstPageNumber).toBe(4);
        expect(setup.printQuality).toBe('600');
        expect(reloaded.getSheetPageSetupXml(0)).toContain('copies="3"');
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
