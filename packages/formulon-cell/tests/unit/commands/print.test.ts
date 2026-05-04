import { describe, expect, it, vi } from 'vitest';
import {
  buildPrintDocument,
  parsePrintTitleCols,
  parsePrintTitleRows,
  printSheet,
} from '../../../src/commands/print.js';
import { WorkbookHandle, addrKey } from '../../../src/engine/workbook-handle.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

/** Seed a single number cell into both the engine and the store, mirroring the
 *  hydrate path that mount.ts uses on cell writes. */
const setNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: number,
): void => {
  wb.setNumber({ sheet: 0, row, col }, value);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

const setText = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: string,
): void => {
  wb.setText({ sheet: 0, row, col }, value);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'text', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

describe('parsePrintTitleRows', () => {
  it('accepts simple "1:3" and "$1:$3" forms', () => {
    expect(parsePrintTitleRows('1:3')).toEqual([0, 2]);
    expect(parsePrintTitleRows('$1:$3')).toEqual([0, 2]);
    expect(parsePrintTitleRows('5')).toEqual([4, 4]);
  });
  it('rejects bad input', () => {
    expect(parsePrintTitleRows('')).toBeNull();
    expect(parsePrintTitleRows(undefined)).toBeNull();
    expect(parsePrintTitleRows('abc')).toBeNull();
  });
});

describe('parsePrintTitleCols', () => {
  it('accepts "A:B" / "$A:$B" forms', () => {
    expect(parsePrintTitleCols('A:B')).toEqual([0, 1]);
    expect(parsePrintTitleCols('$A:$B')).toEqual([0, 1]);
    expect(parsePrintTitleCols('C')).toEqual([2, 2]);
  });
});

describe('buildPrintDocument', () => {
  it('produces a <table> with one <tr> per row and one <td> per col', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 1);
    setNumber(store, wb, 0, 1, 2);
    setNumber(store, wb, 1, 0, 3);
    setNumber(store, wb, 1, 1, 4);

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toContain('<table');
    expect(doc.html).toContain('<tbody>');
    // 2 rows × 2 cols = 4 <td> cells.
    const tdMatches = doc.html.match(/<td/g) ?? [];
    expect(tdMatches.length).toBe(4);
    const trMatches = doc.html.match(/<tr>/g) ?? [];
    // Two body rows, no thead (no print titles, no headings).
    expect(trMatches.length).toBe(2);
    // Cell text rendered.
    expect(doc.html).toContain('>1<');
    expect(doc.html).toContain('>4<');
    wb.dispose();
  });

  it('produces colspan/rowspan for merges', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setText(store, wb, 0, 0, 'Header');
    setNumber(store, wb, 1, 0, 10);
    setNumber(store, wb, 1, 1, 20);
    // Merge A1:B1 — anchor (0,0) absorbs (0,1).
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toContain('colspan="2"');
    // The non-anchor merged cell must NOT emit a separate <td>.
    const headerOccurrences = (doc.html.match(/Header/g) ?? []).length;
    expect(headerOccurrences).toBe(1);
    wb.dispose();
  });

  it('reflects format slice (bold, fill) inline on the cell', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 42);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { bold: true, fill: '#ffeecc', color: '#222222' },
    );

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toMatch(/font-weight:600/);
    expect(doc.html).toMatch(/background:#ffeecc/);
    expect(doc.html).toMatch(/color:#222222/);
    wb.dispose();
  });

  it('repeats print-title rows inside <thead>', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setText(store, wb, 0, 0, 'Title');
    setText(store, wb, 1, 0, 'Sub');
    setNumber(store, wb, 2, 0, 100);
    mutators.setPageSetup(store, 0, { printTitleRows: '1:2' });

    const doc = buildPrintDocument(wb, store, 0);
    const theadIdx = doc.html.indexOf('<thead>');
    const tbodyIdx = doc.html.indexOf('<tbody>');
    expect(theadIdx).toBeGreaterThan(-1);
    expect(theadIdx).toBeLessThan(tbodyIdx);
    const head = doc.html.slice(theadIdx, tbodyIdx);
    expect(head).toContain('Title');
    expect(head).toContain('Sub');
    // Body rows should NOT re-render the title rows.
    const body = doc.html.slice(tbodyIdx);
    expect(body).not.toContain('Title');
    expect(body).toContain('100');
    wb.dispose();
  });

  it('emits @page CSS reflecting orientation, paper size, and margins', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 1);
    mutators.setPageSetup(store, 0, {
      orientation: 'landscape',
      paperSize: 'letter',
      margins: { top: 1, right: 1.25, bottom: 1, left: 0.5 },
    });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toContain('@page');
    expect(doc.html).toContain('letter landscape');
    expect(doc.html).toContain('1in 1.25in 1in 0.5in');
    expect(doc.cssVars['--fc-print-orient']).toBe('landscape');
    expect(doc.cssVars['--fc-print-paper']).toBe('letter');
    wb.dispose();
  });
});

describe('printSheet', () => {
  it('mounts an iframe and triggers print() once the document loads', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 1);

    const host = document.createElement('div');
    document.body.appendChild(host);

    // happy-dom's window doesn't ship `print` — assign a stub on the host
    // window so the synchronous fallback path inside printSheet has a target.
    const winPrint = vi.fn();
    const original = (window as unknown as { print?: () => void }).print;
    (window as unknown as { print: () => void }).print = winPrint;

    const remove = printSheet(wb, store, 0, host);
    const iframe = host.querySelector('iframe');
    expect(iframe).not.toBeNull();
    if (iframe?.contentWindow) {
      // Stub the iframe's contentWindow.print too — without it the iframe
      // load handler would try to drive the real print pipeline.
      (iframe.contentWindow as unknown as { print: () => void }).print = vi.fn();
    }
    // Force the load handler to fire (happy-dom queues it asynchronously).
    iframe?.dispatchEvent(new Event('load'));
    // Either the iframe contentWindow.print or the window.print fallback
    // must have fired — assert the contentWindow stub took the call.
    const cwPrint = iframe?.contentWindow as unknown as { print: ReturnType<typeof vi.fn> };
    expect(cwPrint.print).toHaveBeenCalled();
    remove();
    expect(host.querySelector('iframe')).toBeNull();
    if (original) (window as unknown as { print: () => void }).print = original;
    host.remove();
    wb.dispose();
  });
});
