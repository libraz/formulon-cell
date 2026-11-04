import { afterEach, describe, expect, it } from 'vitest';
import { pasteTSV } from '../../../../src/commands/clipboard/paste.js';
import { resetPasteDelimiters } from '../../../../src/commands/clipboard/paste-delimiters.js';
import { parseTSV } from '../../../../src/commands/clipboard/tsv.js';
import { textToColumns } from '../../../../src/commands/text-to-columns.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

afterEach(() => resetPasteDelimiters());

/** Run Text to Columns over a throwaway cell, as the dialog does. */
const runTextToColumns = async (delimiters: string[]) => {
  const wb = await WorkbookHandle.createDefault();
  const store = createSpreadsheetStore();
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'x,y');
  mutators.replaceCells(store, wb.cells(0));
  textToColumns(store.getState(), store, wb, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, delimiters);
  wb.dispose();
};

describe('plain-text paste splits on the last Text to Columns delimiters', () => {
  it('splits on tab alone in a fresh session', () => {
    expect(parseTSV('p,q')).toEqual([['p,q']]);
    expect(parseTSV('r\ts')).toEqual([['r', 's']]);
  });

  it('after a comma run, splits on comma and no longer on tab', async () => {
    await runTextToColumns([',']);
    expect(parseTSV('p,q')).toEqual([['p', 'q']]);
    expect(parseTSV('r\ts')).toEqual([['r\ts']]);
  });

  it('keeps the setting across workbooks: it belongs to the session', async () => {
    await runTextToColumns([',']);
    const other = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    pasteTSV(store.getState(), store, other, '1,234');
    expect(other.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 1 });
    expect(other.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 234 });
    other.dispose();
  });

  it('a tab + space run splits on both; a tab-only run restores the default', async () => {
    await runTextToColumns(['\t', ' ']);
    expect(parseTSV('g h\ti')).toEqual([['g', 'h', 'i']]);
    await runTextToColumns(['\t']);
    expect(parseTSV('g h\ti')).toEqual([['g h', 'i']]);
    expect(parseTSV('p,q')).toEqual([['p,q']]);
  });
});
