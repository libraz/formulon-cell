import { describe, expect, it, vi } from 'vitest';
import {
  buildPrintDocument,
  parsePrintArea,
  parsePrintTitleCols,
  parsePrintTitleRows,
  printSheet,
} from '../../../src/commands/print.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
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

const setError = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  text: string,
): void => {
  wb.setFormula({ sheet: 0, row, col }, '=1/0');
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'error', code: 0, text },
      formula: '=1/0',
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

describe('parsePrintArea', () => {
  it('accepts A1-style rectangular areas', () => {
    expect(parsePrintArea('A1:C3')).toEqual({ row0: 0, col0: 0, row1: 2, col1: 2 });
    expect(parsePrintArea('$B$2:$D$5')).toEqual({ row0: 1, col0: 1, row1: 4, col1: 3 });
    expect(parsePrintArea('C4')).toEqual({ row0: 3, col0: 2, row1: 3, col1: 2 });
  });

  it('rejects malformed areas', () => {
    expect(parsePrintArea('1:3')).toBeNull();
    expect(parsePrintArea('A:C')).toBeNull();
    expect(parsePrintArea('')).toBeNull();
  });
});

describe('buildPrintDocument', () => {
  it('uses an escaped localized document title when provided', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    const doc = buildPrintDocument(wb, store, 0, '印刷 <preview>');

    expect(doc.html).toContain('<title>印刷 &lt;preview&gt;</title>');
    wb.dispose();
  });

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

  it('emits fill pattern backgrounds for print', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 42);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        fill: '#ffeecc',
        fillPattern: 'diagonalDown',
        fillPatternColor: '#336699',
      },
    );

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toMatch(/background:#ffeecc/);
    expect(doc.html).toMatch(/background-image:repeating-linear-gradient\(45deg, #336699/);
    wb.dispose();
  });

  it('emits red color for red negative number formats in print', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, -42);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 0, negativeStyle: 'red' },
      },
    );

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toMatch(/color:#c00000/);
    wb.dispose();
  });

  it('honors the Sheet tab cell error print mode', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setError(store, wb, 0, 0, '#ERR!');

    mutators.setPageSetup(store, 0, { cellErrorsAs: 'displayed' });
    expect(buildPrintDocument(wb, store, 0).html).toContain('#ERR!');

    mutators.setPageSetup(store, 0, { cellErrorsAs: 'blank' });
    expect(buildPrintDocument(wb, store, 0).html).toContain('<td></td>');

    mutators.setPageSetup(store, 0, { cellErrorsAs: 'dash' });
    expect(buildPrintDocument(wb, store, 0).html).toContain('>--<');

    mutators.setPageSetup(store, 0, { cellErrorsAs: 'na' });
    expect(buildPrintDocument(wb, store, 0).html).toContain('#N/A');

    store.setState((s) => ({ ...s, ui: { ...s.ui, showFormulas: true } }));
    expect(buildPrintDocument(wb, store, 0).html).toContain('=1/0');

    wb.setFormula({ sheet: 0, row: 2, col: 2 }, '=A1');
    store.setState((s) => {
      const map = new Map(s.data.cells);
      map.set(addrKey({ sheet: 0, row: 2, col: 2 }), {
        value: { kind: 'number', value: 0 },
        formula: '=A1',
      });
      return { ...s, data: { ...s.data, cells: map } };
    });
    store.setState((s) => ({ ...s, ui: { ...s.ui, r1c1: true } }));
    expect(buildPrintDocument(wb, store, 0).html).toContain('=R[-2]C[-2]');
    wb.dispose();
  });

  it('emits text direction for print', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setText(store, wb, 0, 0, 'hello');
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { textDirection: 'rtl' });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toMatch(/direction:rtl/);
    wb.dispose();
  });

  it('maps extended Excel alignment choices to printable CSS fallbacks', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setText(store, wb, 0, 0, 'hello');
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        align: 'centerContinuous',
        vAlign: 'distributed',
      },
    );
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { align: 'centerContinuous' });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toMatch(/colspan="2"/);
    expect(doc.html).toMatch(/text-align:center/);
    expect(doc.html).toMatch(/vertical-align:middle/);
    wb.dispose();
  });

  it('limits output to the configured print area and can print in black and white', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setText(store, wb, 0, 0, 'outside');
    setText(store, wb, 1, 1, 'inside');
    setText(store, wb, 1, 2, 'also inside');
    setText(store, wb, 2, 3, 'outside col');
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      { fill: '#ffeecc', color: '#cc0000' },
    );
    mutators.setPageSetup(store, 0, {
      printArea: 'B2:C2',
      blackAndWhite: true,
      draftQuality: true,
    });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toContain('inside');
    expect(doc.html).toContain('also inside');
    expect(doc.html).not.toContain('outside');
    expect(doc.html).not.toContain('#ffeecc');
    expect(doc.html).not.toContain('#cc0000');
    expect(doc.html).toContain('fc-print--black-white');
    expect(doc.html).toContain('fc-print--draft');
    wb.dispose();
  });

  it('prints comments as displayed or at the end of the sheet', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setText(store, wb, 0, 0, 'A');
    setText(store, wb, 1, 1, 'B');
    setText(store, wb, 2, 2, 'C');
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'First <note>' });
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 1 }, { comment: 'Second note' });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { comment: 'Outside area' });

    mutators.setPageSetup(store, 0, {
      comments: 'asDisplayed',
      printArea: 'A1:B2',
    });
    const displayed = buildPrintDocument(wb, store, 0).html;
    expect(displayed).toContain('fc-print__comment-note');
    expect(displayed).toContain('First &lt;note&gt;');
    expect(displayed).toContain('Second note');
    expect(displayed).not.toContain('Outside area');
    expect(displayed).not.toContain('Comments and Notes');

    mutators.setPageSetup(store, 0, { comments: 'endOfSheet' });
    const endOfSheet = buildPrintDocument(wb, store, 0).html;
    expect(endOfSheet).toContain('<section class="fc-print__comments">');
    expect(endOfSheet).toContain('<th>A1</th>');
    expect(endOfSheet).toContain('<th>B2</th>');
    expect(endOfSheet).toContain('First &lt;note&gt;');
    expect(endOfSheet).toContain('Second note');
    expect(endOfSheet).not.toContain('Outside area');
    expect(endOfSheet).not.toContain('<div class="fc-print__comment-note">');

    mutators.setPageSetup(store, 0, { comments: 'none' });
    const none = buildPrintDocument(wb, store, 0).html;
    expect(none).not.toContain('First &lt;note&gt;');
    expect(none).not.toContain('<section class="fc-print__comments">');
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

  it('marks manual row page breaks in the print document', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 1);
    setNumber(store, wb, 1, 0, 2);
    mutators.setPageSetup(store, 0, { manualPageBreakRows: [1] });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toContain('<tr class="fc-print__manual-break-row">');
    expect(doc.html).toContain('break-before: page');
    wb.dispose();
  });

  it('marks manual column page breaks in the print document', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 1);
    setNumber(store, wb, 0, 1, 2);
    mutators.setPageSetup(store, 0, {
      manualPageBreakCols: [1],
      showHeadings: true,
    });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toContain('<th class="fc-print__colhead fc-print__manual-break-col">B</th>');
    expect(doc.html).toContain('<td class="fc-print__manual-break-col">2</td>');
    expect(doc.html).toContain('.fc-print__manual-break-col { break-before: page');
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
      headerMargin: 0.2,
      footerMargin: 0.4,
      centerHorizontally: true,
      centerVertically: true,
      scale: 0.75,
      scaleHeaderFooterWithDocument: false,
      alignHeaderFooterWithMargins: false,
      printQuality: '600',
      firstPageNumber: 3,
    });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.html).toContain('@page');
    expect(doc.html).toContain('letter landscape');
    expect(doc.html).toContain('1in 1.25in 1in 0.5in');
    expect(doc.cssVars['--fc-print-orient']).toBe('landscape');
    expect(doc.cssVars['--fc-print-paper']).toBe('letter');
    expect(doc.cssVars['--fc-print-header-margin']).toBe('0.2in');
    expect(doc.cssVars['--fc-print-footer-margin']).toBe('0.4in');
    expect(doc.cssVars['--fc-print-quality']).toBe('600');
    expect(doc.cssVars['--fc-print-first-page']).toBe('3');
    expect(doc.cssVars['--fc-print-scale']).toBe('0.75');
    expect(doc.cssVars['--fc-print-header-footer-scale']).toBe('1');
    expect(doc.html).toContain('counter-reset: page 2');
    expect(doc.html).toContain('class="fc-print--center-h fc-print--center-v fc-print--hf-free"');
    wb.dispose();
  });

  it('lets fit-to-pages override explicit print scaling', async () => {
    const wb = await newWb();
    const store = createSpreadsheetStore();
    setNumber(store, wb, 0, 0, 1);
    mutators.setPageSetup(store, 0, {
      scale: 0.5,
      fitWidth: 1,
      fitHeight: 2,
    });

    const doc = buildPrintDocument(wb, store, 0);
    expect(doc.cssVars['--fc-print-scale']).toBe('1');
    expect(doc.cssVars['--fc-print-fit-width']).toBe('1');
    expect(doc.cssVars['--fc-print-fit-height']).toBe('2');
    expect(doc.html).toContain('fc-print--fit-to-pages');
    expect(doc.html).not.toContain('transform:scale(0.5)');
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

    const remove = printSheet(wb, store, 0, host, 'PDF', 'pdf');
    const iframe = host.querySelector('iframe');
    expect(iframe).not.toBeNull();
    expect(iframe?.dataset.fcPrintMode).toBe('pdf');
    expect(iframe?.srcdoc).toContain('<title>PDF</title>');
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
