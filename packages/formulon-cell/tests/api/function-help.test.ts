/**
 * Function completion and argument help on the Workbook: the names the formula editor offers,
 * the call the caret sits in and what a function is, in the workbook profile's language.
 */
import { Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';

const live: Workbook[] = [];
afterEach(() => {
  for (const wb of live.splice(0)) wb.dispose();
});

const open = async (locale = 'en'): Promise<Workbook> => {
  const wb = await Workbook.createDefault({ locale });
  live.push(wb);
  return wb;
};

describe('Workbook function help', () => {
  it('offers names starting with the typed text first, then names containing it', async () => {
    const wb = await open();
    const found = wb.suggestFunctions('=1+su', 5, { max: 50 });
    expect(found?.token).toBe('su');
    expect(found?.tokenStart).toBe(3);
    const names = found?.names ?? [];
    expect(names).toContain('SUM');
    const firstInner = names.findIndex((n) => !n.toUpperCase().startsWith('SU'));
    expect(firstInner).toBeGreaterThan(0);
    expect(names.slice(firstInner).every((n) => !n.toUpperCase().startsWith('SU'))).toBe(true);
    expect(wb.suggestFunctions('=su', 3)?.names).toHaveLength(8);
  });

  it('offers nothing outside a function name', async () => {
    const wb = await open();
    expect(wb.suggestFunctions('su', 2)).toBeNull();
    expect(wb.suggestFunctions('="su', 4)).toBeNull();
    expect(wb.suggestFunctions('=SUM', 2)).toBeNull();
  });

  it('reports the innermost call and the argument the caret is in', async () => {
    const wb = await open();
    expect(wb.functionCallAt('=SUM(A1,', 8)).toMatchObject({ name: 'SUM', argIndex: 1 });
    const nested = '=IF(SUM(A1,B1)>0,12,ROUND(';
    expect(wb.functionCallAt(nested, 11)).toMatchObject({ name: 'SUM', argIndex: 1 });
    expect(wb.functionCallAt(nested, 16)).toMatchObject({ name: 'IF', argIndex: 0 });
    expect(wb.functionCallAt(nested, 18)).toMatchObject({ name: 'IF', argIndex: 1 });
    expect(wb.functionCallAt(nested, nested.length)).toMatchObject({ name: 'ROUND', argIndex: 0 });
    expect(wb.functionCallAt('=A1+1', 5)).toBeNull();
    // A known function without argument labels is still a call; an unknown name is not.
    expect(wb.functionCallAt('=CONVERT(1,', 11)).toEqual({
      name: 'CONVERT',
      args: [],
      argIndex: 1,
    });
    expect(wb.functionCallAt('=NOSUCHFN(1,', 12)).toBeNull();
  });

  it('works in the German profile: localized names and the semicolon separator', async () => {
    const wb = await open('de');
    expect(wb.suggestFunctions('=SUMM', 5, { max: 50 })?.names).toContain('SUMME');
    expect(wb.functionCallAt('=SUMME(A1;', 10)).toMatchObject({ name: 'SUM', argIndex: 1 });
    expect(wb.functionCallAt('=SUMME(A1,5', 11)).toMatchObject({ name: 'SUM', argIndex: 0 });
    const sum = wb.describeFunction('SUMME');
    expect(sum).toMatchObject({ name: 'SUM', localizedName: 'SUMME' });
    expect(sum?.args.length).toBeGreaterThan(0);
  });

  it('works in the Japanese profile, describing a function in Japanese', async () => {
    const wb = await open('ja');
    expect(wb.suggestFunctions('=VLOO', 5)?.names).toContain('VLOOKUP');
    expect(wb.functionCallAt('=VLOOKUP(A1,B:C,', 16)).toMatchObject({
      name: 'VLOOKUP',
      argIndex: 2,
    });
    const description = wb.describeFunction('SUM')?.description ?? '';
    expect(description).not.toBe('');
    expect(description).toMatch(/[぀-ヿ一-鿿]/u);
  });

  it('offers LAMBDA names of the workbook, and of a sheet when given', async () => {
    const wb = await open();
    const second = wb.addSheet('Second');
    wb.setDefinedName({ name: 'Twice', formula: '=LAMBDA(x,x*2)' });
    wb.setDefinedName({ name: 'Thrice', formula: '=LAMBDA(x,x*3)', scope: 1 });
    expect(wb.suggestFunctions('=T', 2, { max: 50 })?.names).toContain('Twice');
    expect(wb.suggestFunctions('=T', 2, { max: 50 })?.names).not.toContain('Thrice');
    expect(wb.suggestFunctions('=T', 2, { max: 50, sheet: second })?.names).toContain('Thrice');
    expect(wb.suggestFunctions('=T', 2, { max: 50, sheet: 0 })?.names).not.toContain('Thrice');
    expect(wb.functionCallAt('=Twice(', 7)).toMatchObject({ args: ['x'], argIndex: 0 });
    expect(wb.functionCallAt('=Thrice(', 8)).toBeNull();
    expect(wb.functionCallAt('=Thrice(', 8, { sheet: 'Second' })).toMatchObject({ args: ['x'] });
    expect(wb.describeFunction('twice')).toEqual({
      name: 'TWICE',
      localizedName: 'TWICE',
      args: ['x'],
      description: '',
    });
  });

  it('describes a built-in by its canonical name, and nothing else', async () => {
    const wb = await open();
    const sum = wb.describeFunction('sum');
    expect(sum?.name).toBe('SUM');
    expect(sum?.localizedName).toBe('SUM');
    expect(sum?.args.length).toBeGreaterThan(0);
    expect(sum?.description).not.toBe('');
    expect(wb.describeFunction('NOSUCHFUNCTION')).toBeNull();
  });
});
