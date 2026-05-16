import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearPrintArea,
  clearPrintTitles,
  insertManualPageBreak,
  listPageSetups,
  pageSetupForSheet,
  removeManualPageBreak,
  resetManualPageBreaks,
  resetPageSetup,
  setFitToPages,
  setMarginPreset,
  setPageOrientation,
  setPageScale,
  setPageSetup,
  setPaperSize,
  setPrintArea,
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
    expect(setPrintArea(store, 0, '1:3')).toBeNull();

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
