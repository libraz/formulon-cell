import type { Workbook } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const rawOf = (wb: WorkbookHandle): Workbook => (wb as unknown as { wb: Workbook }).wb;

/** Evaluates `formula` in a lone cell beside Table1 (A1:C4, Qty = 20/30/40) and returns the anchor value. */
async function evaluate(formula: string): Promise<unknown> {
  const wb = await WorkbookHandle.createDefault();
  try {
    const sheet = rawOf(wb).sheet(0);
    sheet.setText('A1', 'Item');
    sheet.setText('B1', 'Qty');
    sheet.setText('C1', 'Other');
    for (let r = 2; r <= 4; r++) {
      sheet.setText(`A${r}`, 'x');
      sheet.setNumber(`B${r}`, r * 10);
      sheet.setNumber(`C${r}`, r);
    }
    sheet.addTable({ range: 'A1:C4', name: 'Table1' });
    sheet.setFormula('H1', formula);
    rawOf(wb).recalc();
    return sheet.getValue('H1');
  } finally {
    wb.dispose();
  }
}

describe('a table column reference used as a reference', () => {
  it('counts the data rows with ROWS', async () => {
    expect(await evaluate('=ROWS(Table1[Qty])')).toBe(3);
  });

  it('is a reference to ISREF', async () => {
    expect(await evaluate('=ISREF(Table1[Qty])')).toBe(true);
  });

  it('counts the columns of the whole table with COLUMNS', async () => {
    expect(await evaluate('=COLUMNS(Table1[#All])')).toBe(3);
  });

  it('yields the top data row from ROW', async () => {
    expect(await evaluate('=ROW(Table1[Qty])')).toBe(2);
  });

  it('indexes into the column', async () => {
    expect(await evaluate('=INDEX(Table1[Qty],2)')).toBe(30);
  });

  it('forms a range between two column references', async () => {
    expect(await evaluate('=SUM(Table1[Qty]:Table1[Qty])')).toBe(90);
  });

  it('extends a column reference to a cell through the range operator', async () => {
    expect(await evaluate('=ROWS(Table1[Qty]:C9)')).toBe(8);
  });
});
