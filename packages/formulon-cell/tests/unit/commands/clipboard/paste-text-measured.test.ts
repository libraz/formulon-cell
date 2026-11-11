import { beforeEach, describe, expect, it } from 'vitest';
import { parseHtmlTable, pasteHtml } from '../../../../src/commands/clipboard/html.js';
import { pasteTSV } from '../../../../src/commands/clipboard/paste.js';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { defaultTableOverlay } from '../../../../src/commands/format-as-table.js';
import { addrKey, type WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { newWb, seedAndMirror, setActive } from './paste-special/fixtures.js';

// Expectations are Excel's results through the object model
// (harness/misc_exp.py, misc_exp2.py; evidence/rowcol-shared2).

const value = (wb: WorkbookHandle, row: number, col: number) => wb.getValue({ sheet: 0, row, col });

describe('pasting plain text as Excel measures it', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('applies the number format the typed-looking text implies', () => {
    setActive(store, 0, 0);
    pasteTSV(store.getState(), store, wb, '001\t1,234\t2026-11-03\t12:30\t1e3\t 5 ');
    const numFmt = (col: number) =>
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col }))?.numFmt;
    expect(value(wb, 0, 0)).toEqual({ kind: 'number', value: 1 });
    expect(numFmt(0)).toBeUndefined();
    expect(value(wb, 0, 1)).toEqual({ kind: 'number', value: 1234 });
    expect(numFmt(1)).toEqual({ kind: 'fixed', decimals: 0, thousands: true });
    expect(numFmt(2)).toEqual({ kind: 'date', pattern: 'mm-dd-yy' });
    expect(numFmt(3)).toEqual({ kind: 'time', pattern: 'h:mm' });
    expect(numFmt(4)).toEqual({ kind: 'scientific', decimals: 2 });
    expect(value(wb, 0, 5)).toEqual({ kind: 'number', value: 5 });
    expect(numFmt(5)).toBeUndefined();
  });

  it('keeps an explicit destination format', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 2 },
      },
    );
    setActive(store, 0, 0);
    pasteTSV(store.getState(), store, wb, '1,234');
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.numFmt,
    ).toEqual({
      kind: 'fixed',
      decimals: 2,
    });
  });
});

describe('pasting next to a table as Excel measures it', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  const tableRange = () => store.getState().tables.tables[0]?.range;
  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
    mutators.upsertTableOverlay(
      store,
      defaultTableOverlay('t1', { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }),
    );
  });

  it('extends the table down over a paste in the row below, then right over an adjacent column', () => {
    setActive(store, 3, 0);
    pasteTSV(store.getState(), store, wb, 'a\tb');
    expect(tableRange()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    setActive(store, 4, 0);
    pasteTSV(store.getState(), store, wb, 'c\td');
    expect(tableRange()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
    setActive(store, 0, 2);
    pasteTSV(store.getState(), store, wb, 'h\r\n1\r\n2');
    expect(tableRange()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 });
  });

  it('extends the table for a structured paste too', () => {
    seedAndMirror(store, wb, [
      { row: 10, col: 0, value: 1 },
      { row: 10, col: 1, value: 2 },
    ]);
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 10, c0: 0, r1: 10, c1: 1 });
    if (!snap) throw new Error('no snapshot');
    setActive(store, 3, 0);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    expect(tableRange()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
  });

  it('leaves the table alone for a paste that is not adjacent', () => {
    setActive(store, 5, 0);
    pasteTSV(store.getState(), store, wb, 'a\tb');
    expect(tableRange()).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
  });
});

describe('pasting an HTML table as Excel measures it', () => {
  const html =
    '<table><tr><th colspan=2 style="background:#ff0">H</th></tr>' +
    '<tr><td><b>bold</b></td><td style="color:red">red</td></tr>' +
    '<tr><td rowspan=2>r2</td><td>x</td></tr><tr><td>y</td></tr></table>';

  it('places each cell in reading order, ignoring colspan and rowspan', () => {
    expect(parseHtmlTable(html)?.map((row) => row.map((cell) => cell.text))).toEqual([
      ['H'],
      ['bold', 'red'],
      ['r2', 'x'],
      ['y'],
    ]);
  });

  it('writes the values and keeps bold, with no fill and no merge', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    setActive(store, 0, 0);
    const result = pasteHtml(store.getState(), store, wb, html, null);
    expect(result.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    expect(value(wb, 0, 0)).toEqual({ kind: 'text', value: 'H' });
    expect(value(wb, 3, 0)).toEqual({ kind: 'text', value: 'y' });
    const fmt = (row: number, col: number) =>
      store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));
    expect(fmt(1, 0)?.bold).toBe(true);
    expect(fmt(0, 0)?.fill).toBeUndefined();
    expect(store.getState().merges.byAnchor.size).toBe(0);
  });
});
