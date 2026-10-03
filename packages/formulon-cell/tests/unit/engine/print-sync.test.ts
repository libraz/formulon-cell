import { describe, expect, it } from 'vitest';
import { mergePageSetupFragment } from '../../../src/engine/page-setup-xml.js';
import {
  hydratePageSetupFromEngine,
  joinHeaderFooter,
  splitHeaderFooter,
  syncPageSetupToEngine,
} from '../../../src/engine/print-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, getPageSetup } from '../../../src/store/store.js';

interface RecordedWrites {
  pageSetup?: Record<string, unknown>;
  extras?: Record<string, unknown>;
  margins?: Record<string, number>;
  printOptions?: Record<string, boolean>;
  headerFooter?: Record<string, unknown>;
  printArea?: string;
  printTitles?: [string, string];
  breaks?: { rows: readonly number[]; cols: readonly number[] };
}

/** Engine stand-in: records every write and replays whatever readback the
 *  test seeded, so a hydrate/flush pair can be asserted without WASM. */
function fakeHandle(
  readback: Partial<{
    pageSetup: ReturnType<WorkbookHandle['getSheetPageSetup']>;
    margins: ReturnType<WorkbookHandle['getSheetPageMargins']>;
    printOptions: ReturnType<WorkbookHandle['getSheetPrintOptions']>;
    headerFooter: ReturnType<WorkbookHandle['getSheetHeaderFooter']>;
    extras: ReturnType<WorkbookHandle['getSheetPageSetupExtras']>;
    printArea: string | null;
    printTitles: { repeatRows: string; repeatCols: string } | null;
    breaks: { rows: number[]; cols: number[] } | null;
  }> = {},
  capabilities: { printSettings?: boolean; printSettingsXml?: boolean; pageBreaks?: boolean } = {
    printSettings: true,
    printSettingsXml: true,
    pageBreaks: true,
  },
): { wb: WorkbookHandle; writes: RecordedWrites } {
  const writes: RecordedWrites = {};
  const wb = {
    capabilities,
    withEngineSyncMuted: <T>(fn: () => T): T => fn(),
    getSheetPageSetup: () => readback.pageSetup ?? null,
    getSheetPageMargins: () => readback.margins ?? null,
    getSheetPageSetupExtras: () => readback.extras ?? null,
    getSheetPrintOptions: () => readback.printOptions ?? null,
    getSheetHeaderFooter: () => readback.headerFooter ?? null,
    getSheetPrintArea: () => readback.printArea ?? null,
    getSheetPrintTitles: () => readback.printTitles ?? null,
    getSheetPageBreaks: () => readback.breaks ?? null,
    setSheetPageSetup: (_s: number, v: Record<string, unknown>) => {
      writes.pageSetup = v;
      return true;
    },
    setSheetPageMargins: (_s: number, v: Record<string, number>) => {
      writes.margins = v;
      return true;
    },
    setSheetPageSetupExtras: (_s: number, v: Record<string, unknown>) => {
      writes.extras = v;
      return true;
    },
    setSheetPrintOptions: (_s: number, v: Record<string, boolean>) => {
      writes.printOptions = v;
      return true;
    },
    setSheetHeaderFooter: (_s: number, v: Record<string, unknown>) => {
      writes.headerFooter = v;
      return true;
    },
    setSheetPrintArea: (_s: number, v: string) => {
      writes.printArea = v;
      return true;
    },
    setSheetPrintTitles: (_s: number, rows: string, cols: string) => {
      writes.printTitles = [rows, cols];
      return true;
    },
    setSheetPageBreaks: (_s: number, v: { rows: readonly number[]; cols: readonly number[] }) => {
      writes.breaks = v;
      return true;
    },
  } as unknown as WorkbookHandle;
  return { wb, writes };
}

describe('header/footer sections', () => {
  it('splits the three section markers and treats a leading run as centre', () => {
    expect(splitHeaderFooter('&LLeft&CMiddle&RRight')).toEqual({
      left: 'Left',
      center: 'Middle',
      right: 'Right',
    });
    expect(splitHeaderFooter('Plain')).toEqual({ left: '', center: 'Plain', right: '' });
  });

  it('keeps a non-section format code with the text it decorates', () => {
    expect(splitHeaderFooter('&CPage &P of &N')).toEqual({
      left: '',
      center: 'Page &P of &N',
      right: '',
    });
    expect(splitHeaderFooter('&L&"MS Gothic"Report')).toEqual({
      left: '&"MS Gothic"Report',
      center: '',
      right: '',
    });
  });

  it('does not read an escaped ampersand as a section marker', () => {
    // `&&L` is a literal ampersand followed by an L, not the left section.
    expect(splitHeaderFooter('&C&&Lorem')).toEqual({ left: '', center: '&&Lorem', right: '' });
  });

  it('round-trips through join, and joins nothing to the empty string', () => {
    const sections = { left: 'A', center: 'B', right: 'C' };
    expect(splitHeaderFooter(joinHeaderFooter(sections))).toEqual(sections);
    expect(joinHeaderFooter({ left: '', center: '', right: '' })).toBe('');
    expect(joinHeaderFooter({ center: 'only' })).toBe('&Conly');
  });
});

describe('hydratePageSetupFromEngine', () => {
  it('maps the engine record onto the store slice', () => {
    const store = createSpreadsheetStore();
    const { wb } = fakeHandle({
      pageSetup: {
        orientation: 2,
        paperSize: 1,
        scale: 75,
        fitToWidth: 0,
        fitToHeight: 0,
        fitToPage: false,
        orientationStated: true,
        paperSizeStated: true,
        scaleStated: true,
        fitToPageStated: true,
      },
      margins: { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.4, footer: 0.4 },
      printOptions: {
        gridLines: true,
        headings: false,
        horizontalCentered: true,
        verticalCentered: false,
      },
      headerFooter: {
        oddHeader: '&CQuarterly report',
        oddFooter: '&RPage &P',
        differentOddEven: false,
        differentFirst: true,
        scaleWithDoc: false,
        alignWithMargins: true,
      },
      printArea: 'A1:F20',
      printTitles: { repeatRows: '1:2', repeatCols: '' },
      breaks: { rows: [24, 48], cols: [] },
    });

    hydratePageSetupFromEngine(wb, store, 0);
    const setup = getPageSetup(store.getState(), 0);

    expect(setup.orientation).toBe('landscape');
    expect(setup.paperSize).toBe('letter');
    expect(setup.paperSizeCode).toBeUndefined();
    expect(setup.scale).toBeCloseTo(0.75);
    expect(setup.margins).toEqual({ left: 0.5, right: 0.5, top: 1, bottom: 1 });
    expect(setup.headerMargin).toBe(0.4);
    expect(setup.showGridlines).toBe(true);
    expect(setup.showHeadings).toBe(false);
    expect(setup.centerHorizontally).toBe(true);
    expect(setup.headerCenter).toBe('Quarterly report');
    expect(setup.footerRight).toBe('Page &P');
    expect(setup.differentFirstPage).toBe(true);
    expect(setup.scaleHeaderFooterWithDocument).toBe(false);
    expect(setup.printArea).toBe('A1:F20');
    expect(setup.printTitleRows).toBe('1:2');
    expect(setup.printTitleCols).toBeUndefined();
    expect(setup.manualPageBreakRows).toEqual([24, 48]);
    expect(setup.manualPageBreakCols).toBeUndefined();
  });

  it('keeps an unmodelled paper code so a save does not re-paper the sheet', () => {
    const store = createSpreadsheetStore();
    // 12 is B4 — outside the sizes the store's `PaperSize` union names.
    const { wb, writes } = fakeHandle({
      pageSetup: {
        orientation: 0,
        paperSize: 12,
        scale: 100,
        fitToWidth: 0,
        fitToHeight: 0,
        fitToPage: false,
        orientationStated: false,
        paperSizeStated: true,
        scaleStated: false,
        fitToPageStated: false,
      },
    });

    hydratePageSetupFromEngine(wb, store, 0);
    expect(getPageSetup(store.getState(), 0).paperSizeCode).toBe(12);

    syncPageSetupToEngine(wb, store, 0);
    expect(writes.pageSetup?.paperSize).toBe(12);
  });

  it('reads fit-to-page back as the store’s width/height constraints', () => {
    const store = createSpreadsheetStore();
    const { wb } = fakeHandle({
      pageSetup: {
        orientation: 1,
        paperSize: 9,
        scale: 100,
        fitToWidth: 1,
        fitToHeight: 0,
        fitToPage: true,
        orientationStated: true,
        paperSizeStated: true,
        scaleStated: false,
        fitToPageStated: true,
      },
    });

    hydratePageSetupFromEngine(wb, store, 0);
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.fitWidth).toBe(1);
    // Zero means "automatic" on that axis, not "fit onto zero pages".
    expect(setup.fitHeight).toBeUndefined();
  });

  it('leaves the slice alone without the engine capability', () => {
    const store = createSpreadsheetStore();
    const { wb } = fakeHandle({ printArea: 'A1:B2' }, { printSettings: false });
    hydratePageSetupFromEngine(wb, store, 0);
    expect(store.getState().pageSetup.setupBySheet.size).toBe(0);
  });
});

describe('the <pageSetup> attributes beside the typed setter', () => {
  it('hydrates the store fields the raw fragment carries', () => {
    const store = createSpreadsheetStore();
    const { wb } = fakeHandle({
      extras: {
        blackAndWhite: true,
        draft: true,
        cellComments: 'atEnd',
        errors: 'NA',
        pageOrder: 'overThenDown',
        firstPageNumber: 7,
        useFirstPageNumber: true,
        horizontalDpi: 600,
        verticalDpi: 600,
      },
    });

    hydratePageSetupFromEngine(wb, store, 0);
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.blackAndWhite).toBe(true);
    expect(setup.draftQuality).toBe(true);
    expect(setup.comments).toBe('endOfSheet');
    expect(setup.cellErrorsAs).toBe('na');
    expect(setup.pageOrder).toBe('overThenDown');
    expect(setup.firstPageNumber).toBe(7);
    expect(setup.printQuality).toBe('600');
  });

  it('ignores a first page number the sheet does not ask to use', () => {
    const store = createSpreadsheetStore();
    const { wb } = fakeHandle({
      extras: {
        blackAndWhite: false,
        draft: false,
        cellComments: '',
        errors: '',
        pageOrder: '',
        firstPageNumber: 5,
        useFirstPageNumber: false,
        horizontalDpi: null,
        verticalDpi: null,
      },
    });

    hydratePageSetupFromEngine(wb, store, 0);
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.firstPageNumber).toBeUndefined();
    expect(setup.comments).toBe('none');
    expect(setup.cellErrorsAs).toBe('displayed');
    expect(setup.pageOrder).toBe('downThenOver');
    expect(setup.printQuality).toBe('automatic');
  });

  it('states a non-default and leaves a schema default unstated', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle();
    store.setState((s) => ({
      ...s,
      pageSetup: {
        setupBySheet: new Map([
          [
            0,
            {
              ...getPageSetup(s, 0),
              blackAndWhite: true,
              comments: 'asDisplayed' as const,
              cellErrorsAs: 'dash' as const,
              pageOrder: 'overThenDown' as const,
              firstPageNumber: 3,
              printQuality: '1200' as const,
            },
          ],
        ]),
      },
    }));

    syncPageSetupToEngine(wb, store, 0);
    expect(writes.extras).toMatchObject({
      blackAndWhite: true,
      draft: false,
      cellComments: 'asDisplayed',
      errors: 'dash',
      pageOrder: 'overThenDown',
      firstPageNumber: 3,
      useFirstPageNumber: true,
      horizontalDpi: 1200,
      verticalDpi: 1200,
    });
  });

  it('clears the attributes a default-valued record implies', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle();
    syncPageSetupToEngine(wb, store, 0);
    expect(writes.extras).toMatchObject({
      blackAndWhite: false,
      draft: false,
      cellComments: '',
      errors: '',
      pageOrder: '',
      firstPageNumber: null,
      useFirstPageNumber: false,
      horizontalDpi: null,
      verticalDpi: null,
    });
  });
});

describe('syncPageSetupToEngine', () => {
  it('states every modelled setting for the sheet', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle();
    store.setState((s) => ({
      ...s,
      pageSetup: {
        setupBySheet: new Map([
          [
            0,
            {
              ...getPageSetup(s, 0),
              orientation: 'landscape' as const,
              paperSize: 'A3' as const,
              scale: 0.5,
              headerCenter: 'Title',
              footerRight: 'Page &P',
              printArea: 'A1:D10',
              printTitleRows: '1:1',
              manualPageBreakRows: [10],
              manualPageBreakCols: [4],
              showGridlines: true,
            },
          ],
        ]),
      },
    }));

    syncPageSetupToEngine(wb, store, 0);

    expect(writes.pageSetup).toMatchObject({ orientation: 2, paperSize: 8, scale: 50 });
    expect(writes.printOptions).toMatchObject({ gridLines: true, headings: false });
    expect(writes.headerFooter).toMatchObject({ oddHeader: '&CTitle', oddFooter: '&RPage &P' });
    expect(writes.printArea).toBe('A1:D10');
    expect(writes.printTitles).toEqual(['1:1', '']);
    expect(writes.breaks).toEqual({ rows: [10], cols: [4] });
  });

  it('turns fit-to-page on only when an axis is constrained', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle();

    syncPageSetupToEngine(wb, store, 0);
    expect(writes.pageSetup).toMatchObject({ fitToPage: false, fitToWidth: 0, fitToHeight: 0 });

    store.setState((s) => ({
      ...s,
      pageSetup: { setupBySheet: new Map([[0, { ...getPageSetup(s, 0), fitWidth: 2 }]]) },
    }));
    syncPageSetupToEngine(wb, store, 0);
    expect(writes.pageSetup).toMatchObject({ fitToPage: true, fitToWidth: 2, fitToHeight: 0 });
  });

  it('drops a scale the engine would reject rather than failing the update', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle();
    store.setState((s) => ({
      ...s,
      pageSetup: { setupBySheet: new Map([[0, { ...getPageSetup(s, 0), scale: Number.NaN }]]) },
    }));
    syncPageSetupToEngine(wb, store, 0);
    expect(writes.pageSetup).not.toHaveProperty('scale');
    expect(writes.pageSetup).toMatchObject({ orientation: 1 });
  });

  it('drops a margin the engine would reject and keeps the rest', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle();
    store.setState((s) => ({
      ...s,
      pageSetup: {
        setupBySheet: new Map([
          [0, { ...getPageSetup(s, 0), margins: { top: -1, right: 0.7, bottom: 0.75, left: 0.7 } }],
        ]),
      },
    }));
    syncPageSetupToEngine(wb, store, 0);
    expect(writes.margins).not.toHaveProperty('top');
    expect(writes.margins).toMatchObject({ right: 0.7, bottom: 0.75, left: 0.7 });
  });

  it('leaves breaks alone when the engine has no break table', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle({}, { printSettings: true, pageBreaks: false });
    syncPageSetupToEngine(wb, store, 0);
    expect(writes.printArea).toBe('');
    expect(writes.breaks).toBeUndefined();
  });

  it('writes nothing without the engine capability', () => {
    const store = createSpreadsheetStore();
    const { wb, writes } = fakeHandle({}, { printSettings: false });
    syncPageSetupToEngine(wb, store, 0);
    expect(writes).toEqual({});
  });
});

describe('mergePageSetupFragment', () => {
  it('adds, replaces and removes only the attributes it is given', () => {
    const merged = mergePageSetupFragment('<pageSetup orientation="landscape" draft="1"/>', {
      blackAndWhite: true,
      draft: false,
    });
    expect(merged).toContain('orientation="landscape"');
    expect(merged).toContain('blackAndWhite="1"');
    expect(merged).not.toContain('draft=');
  });

  it('leaves an attribute this package does not model alone', () => {
    const merged = mergePageSetupFragment('<pageSetup copies="3" horizontalDpi="200"/>', {
      blackAndWhite: true,
    });
    expect(merged).toContain('copies="3"');
    expect(merged).toContain('horizontalDpi="200"');
  });

  it('keeps a printerSettings reference, prefix and all', () => {
    // The engine hands the fragment back with the `r` prefix but no
    // declaration of its own — the worksheet root carried it — so the merge
    // has to parse and re-emit it in exactly that shape.
    const merged = mergePageSetupFragment('<pageSetup r:id="rId1" orientation="portrait"/>', {
      pageOrder: 'overThenDown',
    });
    expect(merged).toContain('r:id="rId1"');
    expect(merged).not.toContain('xmlns:r');
    expect(merged).toContain('pageOrder="overThenDown"');
  });

  it('builds a fragment for a sheet that declares no page setup', () => {
    expect(mergePageSetupFragment('', { blackAndWhite: true })).toContain('blackAndWhite="1"');
  });

  it('reports a fragment it cannot parse rather than replacing it', () => {
    expect(mergePageSetupFragment('<pageSetup', { draft: true })).toBeNull();
    expect(mergePageSetupFragment('<printOptions/>', { draft: true })).toBeNull();
  });
});
