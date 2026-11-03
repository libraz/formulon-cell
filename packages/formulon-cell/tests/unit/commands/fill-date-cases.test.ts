import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { type FillOptions, fillRange } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

interface DateCase {
  seeds: string[];
  format: string;
  type: number | null;
  values: string[];
}

// Captured from desktop Excel (ja) by Range.AutoFill from one or two date
// seeds down to six cells; `type` is the xlAutoFillType (5 days, 6 weekdays,
// 7 months, 8 years).
const cases = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill-dates.json'), 'utf8'),
) as DateCase[];

const OPTIONS: Record<number, FillOptions> = {
  5: { dateUnit: 'days' },
  6: { dateUnit: 'weekdays' },
  7: { dateUnit: 'months' },
  8: { dateUnit: 'years' },
};

const serialOf = (iso: string): number => Date.parse(`${iso}T00:00:00Z`) / 86_400_000 + 25_569;
const isoOf = (serial: number): string =>
  new Date(Math.round((serial - 25_569) * 86_400_000)).toISOString().slice(0, 10);

const TEXT_CASES: [string[], string[]][] = [
  [
    ['1月', '3月'],
    ['1月', '3月', '5月', '7月', '9 月', '11月'],
  ],
  [
    ['火', '木'],
    ['火', '木', '土', '月', '水', '金'],
  ],
  [['令和1年'], ['令和1年', '令和2年', '令和3年', '令和4年', '令和5年', '令和6年']],
  [['平成30年'], ['平成30年', '平成31年', '平成32年', '平成33年', '平成34年', '平成35年']],
  [['1日'], ['1日', '2日', '3日', '4日', '5日', '6日']],
  [['Item1'], ['Item1', 'Item2', 'Item3', 'Item4', 'Item5', 'Item6']],
  [['Jan'], ['Jan', 'Jan', 'Jan', 'Jan', 'Jan', 'Jan']],
];

describe('date autofill matches desktop Excel', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale('ja-JP', 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  let column = 0;
  cases.forEach((c, index) => {
    it(`${c.seeds.join(', ')} ${c.format} type ${c.type ?? 'default'} (#${index})`, () => {
      const col = column++;
      const store = createSpreadsheetStore();
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        const formats = new Map(s.format.formats);
        c.seeds.forEach((iso, i) => {
          const addr = { sheet: 0, row: i, col };
          const value = serialOf(iso);
          wb.setNumber(addr, value);
          cells.set(addrKey(addr), { value: { kind: 'number', value }, formula: null });
          formats.set(addrKey(addr), { numFmt: { kind: 'date', pattern: c.format } });
        });
        return { ...s, data: { ...s.data, cells }, format: { ...s.format, formats } };
      });
      fillRange(
        store.getState(),
        wb,
        { sheet: 0, r0: 0, c0: col, r1: c.seeds.length - 1, c1: col },
        { sheet: 0, r0: 0, c0: col, r1: 5, c1: col },
        { ...(c.type === null ? {} : OPTIONS[c.type]), formatting: 'with', store },
      );
      const got = Array.from({ length: 6 }, (_, i) => {
        const v = wb.getValue({ sheet: 0, row: i, col });
        return v.kind === 'number' ? isoOf(v.value) : v.kind;
      });
      expect(got).toEqual(c.values);
      for (let i = 0; i < 6; i += 1) {
        const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: i, col }))?.numFmt;
        expect(fmt, `format ${i}`).toEqual({ kind: 'date', pattern: c.format });
      }
    });
  });

  it.each(TEXT_CASES)('text seeds %j fill as Excel does', (seeds, want) => {
    const col = column++;
    const store = createSpreadsheetStore();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      seeds.forEach((text, i) => {
        const addr = { sheet: 0, row: i, col };
        wb.setText(addr, text);
        cells.set(addrKey(addr), { value: { kind: 'text', value: text }, formula: null });
      });
      return { ...s, data: { ...s.data, cells } };
    });
    fillRange(
      store.getState(),
      wb,
      { sheet: 0, r0: 0, c0: col, r1: seeds.length - 1, c1: col },
      { sheet: 0, r0: 0, c0: col, r1: 5, c1: col },
      { formatting: 'with', store },
    );
    const got = Array.from({ length: 6 }, (_, i) => {
      const v = wb.getValue({ sheet: 0, row: i, col });
      return v.kind === 'text' ? v.value : v.kind;
    });
    expect(got).toEqual(want);
  });
});
