import { CalcMode, StatusCode } from '@libraz/formulon';
import { mount, type SpreadsheetInstance, Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { expectFormulonError, makeHost } from './helpers.js';

const open: Array<{ dispose(): void }> = [];

const createWorkbook = async (): Promise<Workbook> => {
  const wb = await Workbook.createDefault({ locale: 'en' });
  open.push(wb);
  return wb;
};

afterEach(() => {
  for (const item of open.splice(0).reverse()) item.dispose();
});

describe('workbook calculation settings', () => {
  it('leaves edits made under Manual dirty and settles them when Auto returns', async () => {
    const wb = await createWorkbook();
    const sheet = wb.sheet(0);
    sheet.setNumber('A1', 1);
    sheet.setFormula('B1', '=A1*2');
    const parts: unknown[] = [];
    wb.on('change', (e) => parts.push(...e.parts));

    wb.setCalcMode(CalcMode.Manual);
    expect(wb.getCalcMode()).toBe(CalcMode.Manual);
    sheet.setNumber('A1', 5);
    expect(sheet.getValue('B1')).toBe(2);

    wb.setCalcMode(CalcMode.Auto);
    expect(sheet.getValue('B1')).toBe(10);
    expect(parts.filter((part) => part === 'calc')).toHaveLength(2);
  });

  it('throws the engine error for a mode it refuses and keeps the mode', async () => {
    const wb = await createWorkbook();
    expectFormulonError(() => wb.setCalcMode(9 as CalcMode), StatusCode.InvalidArgument);
    expect(wb.getCalcMode()).toBe(CalcMode.Auto);
  });

  it('reads back the iterative settings it writes', async () => {
    const wb = await createWorkbook();
    wb.updateIterative({ enabled: true, maxIterations: 50, maxChange: 0.01 });
    expect(wb.getIterative()).toMatchObject({ enabled: true, maxIterations: 50, maxChange: 0.01 });
  });

  it('lists the built-in functions in ascending order and reads the profile', async () => {
    const wb = await createWorkbook();
    const names = wb.listFunctionNames();
    expect(names).toContain('SUM');
    expect([...names].sort()).toEqual([...names]);
    expect(typeof wb.getProfileId()).toBe('string');
  });

  it('re-projects the mounted chrome when the calc mode changes through the workbook', async () => {
    const instance: SpreadsheetInstance = await mount(makeHost(), { locale: 'en' });
    open.push(instance);
    const { store } = spreadsheetInternals(instance);
    let updates = 0;
    const off = store.subscribe(() => {
      updates += 1;
    });
    instance.workbook.setCalcMode(CalcMode.Manual);
    off();
    expect(updates).toBeGreaterThan(0);
    expect(instance.workbook.getCalcMode()).toBe(CalcMode.Manual);
  });
});
