import { type ReplaceCellsOptions, Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';

const open: Workbook[] = [];

const createWorkbook = async (): Promise<Workbook> => {
  const wb = await Workbook.createDefault({ locale: 'en' });
  open.push(wb);
  return wb;
};

afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

describe('sheet find', () => {
  it('finds displayed values by rows or columns, folding case and width', async () => {
    const wb = await createWorkbook();
    const sheet = wb.sheet(0);
    sheet.setText('B1', 'Apple');
    sheet.setText('A2', 'ＡＰＰＬＥ pie');
    sheet.setText('A3', 'pear');
    expect(sheet.findAll({ query: 'apple' })).toEqual([
      { row: 0, col: 1 },
      { row: 1, col: 0 },
    ]);
    expect(sheet.findAll({ query: 'apple', searchBy: 'columns' })).toEqual([
      { row: 1, col: 0 },
      { row: 0, col: 1 },
    ]);
    expect(sheet.findAll({ query: 'Apple', matchCase: true })).toEqual([{ row: 0, col: 1 }]);
    expect(sheet.findAll({ query: 'apple', matchEntireCell: true })).toEqual([{ row: 0, col: 1 }]);
    expect(sheet.findAll({ query: 'p?ar' })).toEqual([{ row: 2, col: 0 }]);
    expect(sheet.findAll({ query: 'apple', range: 'A1:A3' })).toEqual([{ row: 1, col: 0 }]);
  });

  it('looks in formulas and notes', async () => {
    const wb = await createWorkbook();
    const sheet = wb.sheet(0);
    sheet.setNumber('A1', 2);
    sheet.setFormula('B1', '=SUM(A1)');
    sheet.setComment('C3', { text: 'check SUM later' });
    expect(sheet.findAll({ query: 'sum' })).toEqual([]);
    expect(sheet.findAll({ query: 'sum', lookIn: 'formulas' })).toEqual([{ row: 0, col: 1 }]);
    expect(sheet.findAll({ query: 'sum', lookIn: 'notes' })).toEqual([{ row: 2, col: 2 }]);
  });
});

describe('sheet replace', () => {
  it('throws for a missing replacement', async () => {
    const wb = await createWorkbook();
    const options = { query: 'a' } as unknown as ReplaceCellsOptions;
    expect(() => wb.sheet(0).replaceAll(options)).toThrow(TypeError);
  });

  it('replaces in constants and leaves formula cells alone in values mode, as one change', async () => {
    const wb = await createWorkbook();
    const sheet = wb.sheet(0);
    sheet.setText('A1', 'red apple');
    sheet.setText('A2', 'apple');
    sheet.setFormula('A3', '="apple"');
    let changes = 0;
    wb.on('change', () => {
      changes += 1;
    });
    const result = sheet.replaceAll({ query: 'apple', replacement: 'pear' });
    expect(result.status).toBe('applied');
    expect(sheet.getValue('A1')).toBe('red pear');
    expect(sheet.getValue('A2')).toBe('pear');
    expect(sheet.getFormula('A3')).toBe('="apple"');
    expect(changes).toBe(1);
  });

  it('replaces formula text in formulas mode, and refuses a broken formula writing nothing', async () => {
    const wb = await createWorkbook();
    const sheet = wb.sheet(0);
    sheet.setNumber('A1', 3);
    sheet.setFormula('B1', '=SUM(A1)');
    sheet.setFormula('B2', '=SUM(A1)+1');
    sheet.replaceAll({ query: 'SUM', replacement: 'MAX', lookIn: 'formulas' });
    expect(sheet.getFormula('B1')).toBe('=MAX(A1)');
    const refused = sheet.replaceAll({ query: 'A1)', replacement: '', lookIn: 'formulas' });
    expect(refused.status).toBe('refused');
    expect(refused.applied).toEqual([]);
    expect(refused.rejections).toEqual([
      expect.objectContaining({ sheet: 0, at: { row: 0, col: 1 }, code: 'invalid' }),
    ]);
    expect(sheet.getFormula('B1')).toBe('=MAX(A1)');
    expect(sheet.getFormula('B2')).toBe('=MAX(A1)+1');
  });
});
