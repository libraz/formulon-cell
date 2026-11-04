import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { fillRange } from '../../../src/commands/fill.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type Unit = 'months' | 'years' | 'days';
type Ymd = readonly [number, number, number];

const serial = ([y, m, d]: Ymd): number => Date.UTC(y, m - 1, d) / 86_400_000 + 25569;
const ymd = (n: number): Ymd => {
  const d = new Date(Math.round((n - 25569) * 86_400_000));
  return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
};
const d = (s: string): Ymd => {
  const [y, m, day] = s.split('/').map(Number) as [number, number, number];
  return [y, m, day];
};

interface Row {
  name: string;
  seeds: string[];
  dir: 'up' | 'down';
  unit: Unit;
  /** Filled cells in ascending sheet order. */
  want: string[];
}

const ROWS: Row[] = [
  {
    name: 'A',
    seeds: ['2024/12/31'],
    dir: 'up',
    unit: 'months',
    want: ['2024/6/30', '2024/7/31', '2024/8/31', '2024/9/30', '2024/10/31', '2024/11/30'],
  },
  {
    name: 'U',
    seeds: ['2024/1/30'],
    dir: 'down',
    unit: 'months',
    want: ['2024/2/29', '2024/3/30', '2024/4/30', '2024/5/30', '2024/6/30', '2024/7/30'],
  },
  {
    name: 'B',
    seeds: ['2024/12/31', '2025/1/31'],
    dir: 'up',
    unit: 'months',
    want: ['2024/6/30', '2024/7/31', '2024/8/31', '2024/9/30', '2024/10/31', '2024/11/30'],
  },
  {
    name: 'D',
    seeds: ['2024/12/31', '2025/1/10'],
    dir: 'up',
    unit: 'months',
    want: ['2024/9/30', '2024/10/11', '2024/10/31', '2024/11/10', '2024/11/30', '2024/12/11'],
  },
  {
    name: 'E',
    seeds: ['2024/12/30', '2025/1/10'],
    dir: 'up',
    unit: 'months',
    want: ['2024/9/30', '2024/10/11', '2024/10/30', '2024/11/10', '2024/11/30', '2024/12/11'],
  },
  {
    name: 'L',
    seeds: ['2024/12/20', '2025/1/10'],
    dir: 'up',
    unit: 'months',
    want: ['2024/9/20', '2024/10/11', '2024/10/20', '2024/11/10', '2024/11/20', '2024/12/11'],
  },
  {
    name: 'M',
    seeds: ['2024/12/31', '2025/1/20'],
    dir: 'up',
    unit: 'months',
    want: ['2024/9/30', '2024/10/21', '2024/10/31', '2024/11/20', '2024/11/30', '2024/12/21'],
  },
  {
    name: 'R',
    seeds: ['2024/12/10', '2025/1/31'],
    dir: 'up',
    unit: 'months',
    want: ['2024/9/10', '2024/11/1', '2024/10/10', '2024/12/1', '2024/11/10', '2025/1/1'],
  },
  {
    name: 'AA',
    seeds: ['2025/2/10', '2025/1/31'],
    dir: 'up',
    unit: 'months',
    want: ['2024/11/10', '2024/10/31', '2024/12/10', '2024/11/30', '2025/1/10', '2024/12/31'],
  },
  {
    name: 'N (row 9)',
    seeds: ['2024/2/29', '2024/3/7', '2024/3/14'],
    dir: 'up',
    unit: 'months',
    want: ['2023/12/31', '2024/1/5', '2024/1/12', '2024/1/31', '2024/2/5', '2024/2/12'],
  },
  {
    name: 'P',
    seeds: ['2024/2/28', '2024/3/7'],
    dir: 'up',
    unit: 'months',
    want: ['2023/11/28', '2023/12/6', '2023/12/28', '2024/1/5', '2024/1/28', '2024/2/5'],
  },
  {
    name: 'Y',
    seeds: ['2024/3/30', '2024/4/10'],
    dir: 'up',
    unit: 'months',
    want: ['2023/12/30', '2024/1/10', '2024/1/30', '2024/2/10', '2024/2/29', '2024/3/12'],
  },
  {
    name: 'V',
    seeds: ['2024/1/30', '2024/2/10'],
    dir: 'down',
    unit: 'months',
    want: ['2024/2/29', '2024/3/12', '2024/3/30', '2024/4/10', '2024/4/30', '2024/5/11'],
  },
  {
    name: 'W',
    seeds: ['2023/1/29', '2023/2/10'],
    dir: 'down',
    unit: 'months',
    want: ['2023/2/28', '2023/3/13', '2023/3/29', '2023/4/10', '2023/4/29', '2023/5/11'],
  },
  {
    name: 'Z',
    seeds: ['2024/1/31', '2024/2/10'],
    dir: 'down',
    unit: 'months',
    want: ['2024/2/29', '2024/3/12', '2024/3/31', '2024/4/10', '2024/4/30', '2024/5/11'],
  },
  {
    name: 'I',
    seeds: ['2024/12/31', '2025/1/10'],
    dir: 'down',
    unit: 'months',
    want: ['2025/1/31', '2025/2/10', '2025/2/28', '2025/3/13', '2025/3/31', '2025/4/10'],
  },
  {
    name: 'Q',
    seeds: ['2024/2/29', '2024/3/7', '2024/3/14'],
    dir: 'down',
    unit: 'months',
    want: ['2024/3/31', '2024/4/5', '2024/4/12', '2024/4/30', '2024/5/6', '2024/5/13'],
  },
  {
    name: 'AC',
    seeds: ['2024/12/31', '2025/1/10'],
    dir: 'up',
    unit: 'years',
    want: ['2021/12/31', '2022/1/10', '2022/12/31', '2023/1/10', '2023/12/31', '2024/1/10'],
  },
  {
    name: 'J',
    seeds: ['2024/12/31', '2025/1/10'],
    dir: 'up',
    unit: 'days',
    want: ['2024/11/1', '2024/11/11', '2024/11/21', '2024/12/1', '2024/12/11', '2024/12/21'],
  },
  // Contiguous month series: the block shifts by seed count x spacing.
  {
    name: 'H',
    seeds: ['2024/1/31', '2024/2/29'],
    dir: 'up',
    unit: 'months',
    want: ['2023/7/31', '2023/8/31', '2023/9/30', '2023/10/31', '2023/11/30', '2023/12/31'],
  },
];

describe('multi-seed Months/Years fill matches desktop Excel', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });
  afterAll(() => wb.dispose());

  let lane = 0;
  for (const row of ROWS) {
    const col = lane++;
    it(`${row.name} ${row.seeds.join(', ')} ${row.dir} ${row.unit}`, () => {
      const store = createSpreadsheetStore();
      const k = row.seeds.length;
      const n = row.want.length;
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        const formats = new Map(s.format.formats);
        row.seeds.forEach((seed, i) => {
          const r = (row.dir === 'up' ? n : 0) + i;
          const value = serial(d(seed));
          wb.setNumber({ sheet: 0, row: r, col }, value);
          cells.set(`0:${r}:${col}`, { value: { kind: 'number', value }, formula: null });
          formats.set(`0:${r}:${col}`, { numFmt: { kind: 'date', pattern: 'yyyy/mm/dd' } });
        });
        return { ...s, data: { ...s.data, cells }, format: { ...s.format, formats } };
      });
      wb.recalc();
      const srcR0 = row.dir === 'up' ? n : 0;
      const src = { sheet: 0, r0: srcR0, c0: col, r1: srcR0 + k - 1, c1: col };
      const dest =
        row.dir === 'up'
          ? { sheet: 0, r0: 0, c0: col, r1: n + k - 1, c1: col }
          : { sheet: 0, r0: 0, c0: col, r1: k + n - 1, c1: col };
      fillRange(store.getState(), wb, src, dest, { dateUnit: row.unit });
      wb.recalc();
      const first = row.dir === 'up' ? 0 : k;
      const got = row.want.map((_, i) => {
        const v = wb.getValue({ sheet: 0, row: first + i, col });
        const [y, m, day] = ymd(v.kind === 'number' ? v.value : Number.NaN);
        return `${y}/${m}/${day}`;
      });
      expect(got).toEqual(row.want);
    });
  }
});
