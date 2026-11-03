import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { type FillOptions, fillRange } from '../../../src/commands/fill.js';
import { addrKey, formatA1Cell } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type Src = number | string | { f: string };
type Out = number | string;

interface Row {
  name: string;
  source: Src[];
  /** Filled values per type, including the source cells. */
  default: Out[];
  copy: Out[];
  series: Out[];
  /** Source cells carry a date format. */
  date?: boolean;
  formulas?: { default: string[]; copy: string[]; series: string[] };
}

const rep = (v: Out[], n: number): Out[] =>
  Array.from({ length: n }, (_, i) => v[i % v.length] as Out);
const same = (v: Out, n = 5): Out[] => Array.from({ length: n }, () => v);

// Desktop Excel, Range.autofill on the Japanese build; the destination is five
// cells past a one-cell source and four past a two-cell source.
const ROWS: Row[] = [
  { name: '1', source: [1], default: same(1), copy: same(1), series: [1, 2, 3, 4, 5] },
  {
    name: '1,2',
    source: [1, 2],
    default: [1, 2, 3, 4, 5, 6],
    copy: rep([1, 2], 6),
    series: [1, 2, 3, 4, 5, 6],
  },
  {
    name: '1,3',
    source: [1, 3],
    default: [1, 3, 5, 7, 9, 11],
    copy: rep([1, 3], 6),
    series: [1, 3, 5, 7, 9, 11],
  },
  {
    name: 'Item1',
    source: ['Item1'],
    default: ['Item1', 'Item2', 'Item3', 'Item4', 'Item5'],
    copy: same('Item1'),
    series: ['Item1', 'Item2', 'Item3', 'Item4', 'Item5'],
  },
  { name: 'A', source: ['A'], default: same('A'), copy: same('A'), series: same('A') },
  {
    name: 'Mon',
    source: ['Mon'],
    default: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
    copy: same('Mon'),
    series: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
  },
  {
    name: '3rd',
    source: ['3rd'],
    default: ['3rd', '4th', '5th', '6th', '7th'],
    copy: same('3rd'),
    series: ['3rd', '4th', '5th', '6th', '7th'],
  },
  {
    name: 'Q1',
    source: ['Q1'],
    default: ['Q1', 'Q2', 'Q3', 'Q4', 'Q1'],
    copy: same('Q1'),
    series: ['Q1', 'Q2', 'Q3', 'Q4', 'Q1'],
  },
  {
    name: '1月',
    source: ['1月'],
    default: ['1月', '2月', '3月', '4月', '5月'],
    copy: same('1月'),
    series: ['1月', '2月', '3月', '4月', '5月'],
  },
  {
    name: '1,a',
    source: [1, 'a'],
    default: [1, 'a', 2, 'a', 3, 'a'],
    copy: rep([1, 'a'], 6),
    series: [1, 'a', 2, 'a', 3, 'a'],
  },
  {
    name: '1,=A1+1',
    source: [1, { f: '=A1+1' }],
    default: [1, 2, 2, 3, 3, 4],
    copy: [1, 2, 1, 2, 1, 2],
    series: [1, 2, 2, 3, 3, 4],
    formulas: {
      default: ['', '=A1+1', '', '=A3+1', '', '=A5+1'],
      copy: ['', '=A1+1', '', '=A3+1', '', '=A5+1'],
      series: ['', '=A1+1', '', '=A3+1', '', '=A5+1'],
    },
  },
  {
    name: 'date',
    date: true,
    source: [45000],
    default: [45000, 45001, 45002, 45003, 45004],
    copy: same(45000),
    series: [45000, 45001, 45002, 45003, 45004],
  },
];

const TYPES: [keyof Pick<Row, 'default' | 'copy' | 'series'>, FillOptions][] = [
  ['default', {}],
  ['copy', { copyOnly: true }],
  ['series', { series: true }],
];

describe('fill-handle drag matches desktop Excel (autofill cases)', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale('ja-JP', 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  let col = 0;
  for (const row of ROWS) {
    for (const [type, options] of TYPES) {
      const column = col++;
      it(`${row.name} ${type}`, () => {
        const store = createSpreadsheetStore();
        const k = row.source.length;
        const total = row.default.length;
        store.setState((s) => {
          const cells = new Map(s.data.cells);
          row.source.forEach((src, i) => {
            const addr = { sheet: 0, row: i, col: column };
            if (typeof src === 'number') {
              wb.setNumber(addr, src);
              cells.set(addrKey(addr), { value: { kind: 'number', value: src }, formula: null });
            } else if (typeof src === 'string') {
              wb.setText(addr, src);
              cells.set(addrKey(addr), { value: { kind: 'text', value: src }, formula: null });
            } else {
              const formula = src.f.replace('A1', formatA1Cell(0, column));
              wb.setFormula(addr, formula);
              cells.set(addrKey(addr), { value: { kind: 'number', value: 0 }, formula });
            }
          });
          const formats = new Map(s.format.formats);
          if (row.date) {
            formats.set(addrKey({ sheet: 0, row: 0, col: column }), {
              numFmt: { kind: 'date', pattern: 'yyyy/m/d' },
            });
          }
          return { ...s, data: { ...s.data, cells }, format: { ...s.format, formats } };
        });
        wb.recalc();
        fillRange(
          store.getState(),
          wb,
          { sheet: 0, r0: 0, c0: column, r1: k - 1, c1: column },
          { sheet: 0, r0: 0, c0: column, r1: total - 1, c1: column },
          { ...options, formatting: 'with', store },
        );
        wb.recalc();
        const got = Array.from({ length: total }, (_, i) => {
          const v = wb.getValue({ sheet: 0, row: i, col: column });
          return v.kind === 'number' || v.kind === 'text' ? v.value : v.kind;
        });
        expect(got).toEqual(row[type]);
        if (row.formulas) {
          const letter = formatA1Cell(0, column).replace('1', '');
          const f = Array.from(
            { length: total },
            (_, i) => wb.cellFormula({ sheet: 0, row: i, col: column }) ?? '',
          );
          expect(f).toEqual(row.formulas[type].map((x) => x.replace(/A(\d)/g, `${letter}$1`)));
        }
      });
    }
  }
});
