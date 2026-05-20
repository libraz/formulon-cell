import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  addPrintArea,
  applyPrinterProfileBounds,
  clearPrintArea,
  clearPrintableBounds,
  clearPrintTitles,
  insertManualPageBreak,
  listPageSetups,
  normalizePrinterProfile,
  normalizePrinterProfileId,
  normalizePrinterProfiles,
  pageSetupForSheet,
  removeManualPageBreak,
  resetManualPageBreaks,
  resetPageSetup,
  resolvePrinterProfileBounds,
  setFitToPages,
  setMarginPreset,
  setPageOrientation,
  setPageScale,
  setPageSetup,
  setPaperSize,
  setPrintArea,
  setPrintableBounds,
  setPrintGridlines,
  setPrintHeadings,
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

  it('sets, validates, and clears print area', () => {
    const store = createSpreadsheetStore();

    expect(setPrintArea(store, 0, ' B2:D5 ')).toMatchObject({ printArea: 'B2:D5' });
    expect(setPrintArea(store, 0, '$A$1:$C$3')).toMatchObject({ printArea: '$A$1:$C$3' });
    expect(addPrintArea(store, 0, 'E1:F2')).toMatchObject({ printArea: '$A$1:$C$3,E1:F2' });
    expect(setPrintArea(store, 0, 'A1:B2,D4:E5')).toMatchObject({
      printArea: 'A1:B2,D4:E5',
    });
    expect(setPrintArea(store, 0, '1:3')).toBeNull();
    expect(addPrintArea(store, 0, 'A:C')).toBeNull();

    const cleared = clearPrintArea(store, 0);
    expect(cleared.printArea).toBeUndefined();
  });

  it('sets scale-to-fit width, height, and explicit scale', () => {
    const store = createSpreadsheetStore();

    expect(setFitToPages(store, 0, 'width', 1).fitWidth).toBe(1);
    expect(setFitToPages(store, 0, 'height', 2).fitHeight).toBe(2);
    expect(setFitToPages(store, 0, 'width', 0).fitWidth).toBeUndefined();

    const scaled = setPageScale(store, 0, 1.25);
    expect(scaled.scale).toBe(1.25);
    expect(scaled.fitWidth).toBeUndefined();
    expect(scaled.fitHeight).toBeUndefined();
    expect(setPageScale(store, 0, 9).scale).toBe(4);
    expect(setPageScale(store, 0, 0).scale).toBe(0.1);
  });

  it('sets print gridlines and headings', () => {
    const store = createSpreadsheetStore();

    expect(setPrintGridlines(store, 0, true).showGridlines).toBe(true);
    expect(setPrintHeadings(store, 0, true).showHeadings).toBe(true);
    expect(setPrintGridlines(store, 0, false).showGridlines).toBe(false);
    expect(setPrintHeadings(store, 0, false).showHeadings).toBe(false);
  });

  it('sets and clears printer printable bounds independently from print area', () => {
    const store = createSpreadsheetStore();

    const setup = setPrintableBounds(store, 0, {
      top: 0.25,
      right: 0.2,
      bottom: 0.3,
      left: 0.2,
    });

    expect(setup.printableBounds).toEqual({ top: 0.25, right: 0.2, bottom: 0.3, left: 0.2 });
    expect(setup.printArea).toBeUndefined();

    const normalized = setPrintableBounds(store, 0, {
      top: -1,
      right: Number.NaN,
      bottom: 0.5,
      left: 0.4,
    });
    expect(normalized.printableBounds).toEqual({ top: 0, right: 0, bottom: 0.5, left: 0.4 });

    const cleared = clearPrintableBounds(store, 0);
    expect(cleared.printableBounds).toBeUndefined();
  });

  it('resolves host printer profiles for the active paper and orientation', () => {
    const setup = { ...defaultPageSetup(), paperSize: 'letter' as const };
    setup.orientation = 'landscape';
    const profiles = [
      { name: 'generic', printableBounds: { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 } },
      {
        name: 'letter',
        paperSize: 'letter' as const,
        printableBounds: { top: 0.2, right: 0.2, bottom: 0.2, left: 0.2 },
      },
      {
        name: 'letter landscape',
        paperSize: 'letter' as const,
        orientation: 'landscape' as const,
        printableBounds: { top: 0.25, right: 0.18, bottom: 0.25, left: 0.18 },
      },
    ];

    expect(resolvePrinterProfileBounds(setup, profiles)).toEqual({
      top: 0.25,
      right: 0.18,
      bottom: 0.25,
      left: 0.18,
    });
    expect(
      resolvePrinterProfileBounds(
        setup,
        [
          ...profiles,
          {
            id: 'manual-tray',
            name: 'Manual tray',
            printableBounds: { top: 0.4, right: 0.3, bottom: 0.4, left: 0.3 },
          },
        ],
        ' manual-tray ',
      ),
    ).toEqual({
      top: 0.4,
      right: 0.3,
      bottom: 0.4,
      left: 0.3,
    });
  });

  it('normalizes host printer profiles before profile resolution', () => {
    expect(normalizePrinterProfileId(' office ')).toBe('office');
    expect(normalizePrinterProfileId('   ')).toBeUndefined();
    expect(
      normalizePrinterProfile({
        id: ' office ',
        name: ' Office Printer ',
        paperSize: 'tabloid',
        orientation: 'landscape',
        printableBounds: { top: -1, right: Number.NaN, bottom: 0.3, left: 0.2 },
      }),
    ).toEqual({
      id: 'office',
      name: 'Office Printer',
      paperSize: 'tabloid',
      orientation: 'landscape',
      printableBounds: { top: 0, right: 0, bottom: 0.3, left: 0.2 },
    });
    expect(normalizePrinterProfiles(undefined)).toBeUndefined();
    expect(
      normalizePrinterProfiles([
        { id: 'tray', printableBounds: { top: 0.1 } },
        { id: ' tray ', name: 'Duplicate', printableBounds: { top: 0.2 } },
        { name: 'A3 tray', paperSize: 'A3', printableBounds: { top: 0.3 } },
      ]),
    ).toEqual([
      { id: 'tray', printableBounds: { top: 0.1, right: 0, bottom: 0, left: 0 } },
      {
        name: 'A3 tray',
        paperSize: 'A3',
        printableBounds: { top: 0.3, right: 0, bottom: 0, left: 0 },
      },
    ]);
  });

  it('applies and clears host printer profile bounds on the sheet page setup', () => {
    const store = createSpreadsheetStore();
    setPageSetup(store, 0, { paperSize: 'letter', orientation: 'landscape' });

    const applied = applyPrinterProfileBounds(store, 0, [
      {
        paperSize: 'letter',
        orientation: 'landscape',
        printableBounds: { top: 0.25, right: 0.18, bottom: 0.25, left: 0.18 },
      },
    ]);
    expect(applied.printableBounds).toEqual({
      top: 0.25,
      right: 0.18,
      bottom: 0.25,
      left: 0.18,
    });

    const cleared = applyPrinterProfileBounds(store, 0, [
      { paperSize: 'A4', printableBounds: { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 } },
    ]);
    expect(cleared.printableBounds).toBeUndefined();

    const preferred = applyPrinterProfileBounds(
      store,
      0,
      [
        {
          id: 'default',
          paperSize: 'letter',
          orientation: 'landscape',
          printableBounds: { top: 0.2, right: 0.2, bottom: 0.2, left: 0.2 },
        },
        {
          id: 'manual',
          paperSize: 'letter',
          orientation: 'landscape',
          printableBounds: { top: 0.5, right: 0.4, bottom: 0.5, left: 0.4 },
        },
      ],
      null,
      'manual',
    );
    expect(preferred.printableBounds).toEqual({
      top: 0.5,
      right: 0.4,
      bottom: 0.5,
      left: 0.4,
    });
  });

  it('inserts, removes, and resets manual page breaks', () => {
    const store = createSpreadsheetStore();

    expect(insertManualPageBreak(store, 0, 'row', 10).manualPageBreakRows).toEqual([10]);
    expect(insertManualPageBreak(store, 0, 'row', 5).manualPageBreakRows).toEqual([5, 10]);
    expect(insertManualPageBreak(store, 0, 'row', 10).manualPageBreakRows).toEqual([5, 10]);
    expect(insertManualPageBreak(store, 0, 'col', 3).manualPageBreakCols).toEqual([3]);
    expect(insertManualPageBreak(store, 0, 'row', 0).manualPageBreakRows).toEqual([5, 10]);

    expect(removeManualPageBreak(store, 0, 'row', 5).manualPageBreakRows).toEqual([10]);
    const reset = resetManualPageBreaks(store, 0);
    expect(reset.manualPageBreakRows).toBeUndefined();
    expect(reset.manualPageBreakCols).toBeUndefined();
  });

  it('records page setup commands in history when provided', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    setPageOrientation(store, 0, 'landscape', history);
    expect(pageSetupForSheet(store.getState(), 0).orientation).toBe('landscape');
    expect(history.undo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).orientation).toBe('portrait');
    expect(history.redo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).orientation).toBe('landscape');

    setPaperSize(store, 0, 'letter', history);
    setMarginPreset(store, 0, 'narrow', history);
    setPrintArea(store, 0, 'A1:C3', history);
    expect(pageSetupForSheet(store.getState(), 0)).toMatchObject({
      paperSize: 'letter',
      printArea: 'A1:C3',
    });

    expect(history.undo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).printArea).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).margins).toEqual(defaultPageSetup().margins);
    expect(history.undo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).paperSize).toBe(defaultPageSetup().paperSize);
  });

  it('does not push history entries for invalid print area or title input', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    expect(setPrintArea(store, 0, '1:3', history)).toBeNull();
    expect(setPrintTitleRows(store, 0, 'abc', history)).toBeNull();
    expect(history.canUndo()).toBe(false);
  });

  it('does not push history entries when page setup is unchanged', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    setPrintArea(store, 0, 'A1:C3', history);
    expect(history.undo()).toBe(true);
    expect(history.redo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).printArea).toBe('A1:C3');

    setPrintArea(store, 0, 'A1:C3', history);
    expect(history.undo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).printArea).toBeUndefined();
    expect(history.undo()).toBe(false);

    insertManualPageBreak(store, 0, 'row', 10, history);
    expect(pageSetupForSheet(store.getState(), 0).manualPageBreakRows).toEqual([10]);
    insertManualPageBreak(store, 0, 'row', 10, history);
    expect(history.undo()).toBe(true);
    expect(pageSetupForSheet(store.getState(), 0).manualPageBreakRows).toBeUndefined();
    expect(history.undo()).toBe(false);
  });
});
