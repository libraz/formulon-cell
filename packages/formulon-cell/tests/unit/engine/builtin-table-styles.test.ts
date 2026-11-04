import { describe, expect, it } from 'vitest';
import { BUILTIN_TABLE_STYLES } from '../../../src/commands/format-as-table.js';
import { builtinTableStyleElements } from '../../../src/engine/builtin-table-styles.js';

/** The theme the definitions were read under: clrScheme "Office" (2007). */
const OFFICE_2007 = [
  0xff000000, 0xffffffff, 0xff1f497d, 0xffeeece1, 0xff4f81bd, 0xffc0504d, 0xff9bbb59, 0xff8064a2,
  0xff4bacc6, 0xfff79646,
];

const rgb = (hex: string | undefined): number[] =>
  hex ? [1, 3, 5].map((i) => Number.parseInt(hex.slice(i, i + 2), 16)) : [];

describe('built-in table style elements', () => {
  it('defines every gallery style and nothing else', () => {
    for (const { name } of BUILTIN_TABLE_STYLES) {
      expect(builtinTableStyleElements(name), name).not.toBeNull();
    }
    expect(builtinTableStyleElements('TableStyleMedium29')).toBeNull();
    expect(builtinTableStyleElements('PivotStyleLight16')).toBeNull();
  });

  it('resolves Medium 2 to the colours the desktop app reports under its theme', () => {
    const style = builtinTableStyleElements('TableStyleMedium2', OFFICE_2007);
    expect(rgb(style?.headerRow?.fill)).toEqual([79, 129, 189]);
    expect(rgb(style?.headerRow?.color)).toEqual([255, 255, 255]);
    expect(style?.headerRow?.bold).toBe(true);
    expect(rgb(style?.firstRowStripe?.fill)).toEqual([220, 230, 241]);
    expect(style?.totalRow?.borders?.top).toEqual({ style: 'double', color: '#4f81bd' });
  });

  it('keeps tinted neutrals and dark fills', () => {
    expect(
      rgb(builtinTableStyleElements('TableStyleLight1', OFFICE_2007)?.firstRowStripe?.fill),
    ).toEqual([217, 217, 217]);
    const dark1 = builtinTableStyleElements('TableStyleDark1', OFFICE_2007);
    expect(rgb(dark1?.headerRow?.fill)).toEqual([0, 0, 0]);
    expect(rgb(dark1?.firstRowStripe?.fill)).toEqual([64, 64, 64]);
    expect(dark1?.headerRow?.borders?.bottom).toEqual({ style: 'medium', color: '#ffffff' });
  });

  it('gives Light 8-14 a second stripe and no other style one', () => {
    const withSecond = BUILTIN_TABLE_STYLES.filter(
      ({ name }) => builtinTableStyleElements(name)?.secondRowStripe !== undefined,
    ).map(({ name }) => name);
    expect(withSecond).toEqual([8, 9, 10, 11, 12, 13, 14].map((n) => `TableStyleLight${n}`));
    expect(builtinTableStyleElements('TableStyleLight9', OFFICE_2007)?.secondRowStripe).toEqual({
      borders: { top: { style: 'thin', color: '#4f81bd' } },
    });
  });

  it('follows the workbook theme', () => {
    const green = [...OFFICE_2007];
    green[4] = 0xff00b050;
    expect(builtinTableStyleElements('TableStyleMedium2', green)?.headerRow?.fill).toBe('#00b050');
  });
});
