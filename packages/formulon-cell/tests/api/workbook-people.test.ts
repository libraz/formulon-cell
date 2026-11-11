import { StatusCode } from '@libraz/formulon';
import { Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { expectFormulonError } from './helpers.js';

const open: Workbook[] = [];

const createWorkbook = async (): Promise<Workbook> => {
  const wb = await Workbook.createDefault({ locale: 'en' });
  open.push(wb);
  return wb;
};

afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

describe('sheet copy', () => {
  it('copies a sheet in front of another with its cells and reports the sheet list change', async () => {
    const wb = await createWorkbook();
    const first = wb.sheet(0);
    first.setNumber('A1', 7);
    first.setFormula('B1', '=A1+1');
    wb.addSheet('Last');
    const parts: unknown[] = [];
    wb.on('change', (e) => parts.push(...e.parts));

    const copy = first.copy({ insertBefore: 'Last', newName: 'Copy' });
    expect(copy.name).toBe('Copy');
    expect(copy.index).toBe(1);
    expect(copy.getValue('B1')).toBe(8);
    expect(wb.listSheets().map((s) => s.name)).toEqual([first.name, 'Copy', 'Last']);
    expect(parts).toEqual(['sheets']);
  });

  it('appends a copy named after its source when given nothing', async () => {
    const wb = await createWorkbook();
    const copy = wb.sheet(0).copy();
    expect(copy.index).toBe(1);
    expect(copy.name).toBe(`${wb.sheet(0).name} (2)`);
  });
});

describe('people and external links', () => {
  it('adds, reads and removes a person', async () => {
    const wb = await createWorkbook();
    const parts: unknown[] = [];
    wb.on('change', (e) => parts.push(...e.parts));
    const id = wb.addPerson({ displayName: 'Ada' });
    expect(wb.getPerson(id)).toMatchObject({ id, displayName: 'Ada', userId: null });
    expect(wb.listPeople().map((p) => p.displayName)).toEqual(['Ada']);
    wb.removePerson(id);
    expect(wb.getPerson(id)).toBeNull();
    expect(parts).toEqual(['threadedComments', 'threadedComments']);
  });

  it('refuses a person without a name', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.addPerson({ displayName: '' }), StatusCode.InvalidArgument);
  });

  it('lists the external books formulas name', async () => {
    const wb = await createWorkbook();
    expect(wb.listExternalLinks()).toEqual([]);
    wb.sheet(0).setFormula('A1', '=[Other.xlsx]Sheet1!A1');
    expect(wb.listExternalLinks()).toHaveLength(1);
  });
});
