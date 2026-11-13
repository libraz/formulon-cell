// @vitest-environment node

import { FontScheme } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

// Raw widths Mac Excel saved for 8.43 / 10 / 20 displayed characters in a 游ゴシック 12 book,
// re-read under each Normal font, and the default column with base width 10 (xlwings
// measurements). On a Mac profile the engine lays these out; cell reads its points at 4/3 px.
const SAVED_RAW = [9.140625, 10.7109375, 20.7109375] as const;
const MEASURED: ReadonlyArray<readonly [string, number, readonly number[], number]> = [
  ['Aptos Narrow', 12, [55, 64, 124], 65],
  ['Calibri', 12, [55, 64, 124], 65],
  ['Calibri', 11, [55, 64, 124], 65],
  ['Arial', 10, [55, 64, 124], 65],
  ['ＭＳ Ｐゴシック', 11, [55, 64, 124], 65],
  ['Aptos Narrow', 11, [55, 64, 124], 65],
  ['游ゴシック', 11, [55, 64, 124], 65],
  ['游ゴシック', 12, [64, 75, 145], 75],
];
// Measured with a 7 pt digit, which the engine lays out at 6 pt; each passes once the engine
// carries the font's Mac row.
const ENGINE_DISAGREES: typeof MEASURED = [
  ['Meiryo', 11, [64, 75, 145], 75],
  ['Helvetica Neue', 12, [64, 75, 145], 75],
  ['Arial', 12, [64, 75, 145], 75],
];
const PX_PER_PT = 96 / 72;

const handles: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of handles.splice(0)) wb.dispose();
});

const layOut = async (font: string, size: number) => {
  const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
  handles.push(wb);
  // A named Normal font drops the theme link, or Mac Excel draws the theme's font instead.
  expect(wb.setWorkbookDefaultFont({ name: font, sizePt: size, scheme: FontScheme.None })).toBe(
    true,
  );
  const raw = (wb as unknown as { wb: { sheet(i: number): RawSheet } }).wb.sheet(0);
  for (const [col, chars] of SAVED_RAW.entries()) raw.setColumnWidth(col, col, chars);
  return {
    widths: wb.getColumnLayouts(0).map((c) => Math.round((c.width / PX_PER_PT) * 1e6) / 1e6),
    base: (wb.getSheetDefaultLayoutPx(0)?.defaultColWidth ?? Number.NaN) / PX_PER_PT,
  };
};

describe('Mac column geometry from the engine profile', () => {
  it.each(MEASURED)('lays out %s %d pt columns as measured', async (font, size, widths, base) => {
    const laid = await layOut(font, size);
    expect(laid.widths).toEqual(widths);
    expect(laid.base).toBeCloseTo(base, 6);
  });

  for (const [font, size, widths] of ENGINE_DISAGREES) {
    it.fails(`lays out ${font} ${size} pt columns as measured (engine disagrees)`, async () => {
      expect((await layOut(font, size)).widths).toEqual(widths);
    });
  }
});

type RawSheet = { setColumnWidth(first: number, last: number, chars: number): void };
