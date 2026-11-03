import { describe, expect, it } from 'vitest';
import {
  pageSetupForSheet,
  setFitToPages,
  setPageScale,
  setPageSetup,
  setPrintArea,
  setPrintTitleCols,
  setPrintTitleRows,
} from '../../../src/commands/page-setup.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// Excel PageSetup through the object model;
// margins in inches (Excel points / 72).
describe('page setup as Excel measures it', () => {
  it('starts a new sheet on the Excel defaults', () => {
    const setup = pageSetupForSheet(createSpreadsheetStore().getState(), 0);
    expect(setup).toMatchObject({
      margins: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7 },
      headerMargin: 0.3,
      footerMargin: 0.3,
      scale: 1,
      centerHorizontally: false,
      centerVertically: false,
      showGridlines: false,
      showHeadings: false,
      blackAndWhite: false,
      draftQuality: false,
      pageOrder: 'downThenOver',
    });
    expect(setup.firstPageNumber).toBeUndefined();
    expect(setup.printArea).toBeUndefined();
    expect(setup.printTitleRows).toBeUndefined();
    expect(setup.printTitleCols).toBeUndefined();
  });

  it('round-trips margins, zoom, print area, titles and header/footer codes verbatim', () => {
    const store = createSpreadsheetStore();
    setPageSetup(store, 0, { margins: { left: 0.5 } });
    setPageScale(store, 0, 1.5);
    setPrintArea(store, 0, '$A$1:$C$5');
    setPrintTitleRows(store, 0, '$1:$2');
    setPrintTitleCols(store, 0, '$A:$B');
    setPageSetup(store, 0, { footerLeft: '&P / &N', headerCenter: '&D &T &F' });
    const setup = pageSetupForSheet(store.getState(), 0);
    expect(setup.margins.left).toBe(0.5);
    expect(setup.scale).toBe(1.5);
    expect(setup.printArea).toBe('$A$1:$C$5');
    expect(setup.printTitleRows).toBe('$1:$2');
    expect(setup.printTitleCols).toBe('$A:$B');
    expect(setup.footerLeft).toBe('&P / &N');
    expect(setup.headerCenter).toBe('&D &T &F');
  });

  it('keeps the zoom when fit-to-pages is set', () => {
    const store = createSpreadsheetStore();
    setPageScale(store, 0, 1.5);
    setFitToPages(store, 0, 'width', 1);
    expect(pageSetupForSheet(store.getState(), 0).scale).toBe(1.5);
  });
});
