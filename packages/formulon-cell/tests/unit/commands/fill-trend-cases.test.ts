import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { type FillOptions, fillRange } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type Seed = number | string | null;
type Out = number | string | null;

interface Fixture {
  linear: Array<{ seeds: number[]; type: number; values: number[] }>;
  reverse: Array<{
    name: string;
    dir: 'up' | 'left' | 'right';
    seeds: Seed[];
    type: number | null;
    format: string | null;
    values: Out[];
  }>;
  blocks: Array<{
    name: string;
    grid: Seed[][];
    type: number | null;
    dir: 'down' | 'right';
    values: Out[] | Out[][];
  }>;
}

// Captured from desktop Excel (ja) by Range.AutoFill; `type` is the
// xlAutoFillType (1 copy, 2 series, 6 weekdays, 7 months, 8 years, 9 linear
// trend, 10 growth trend).
const fixture = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill-trend.json'), 'utf8'),
) as Fixture;

const OPTIONS: Record<number, FillOptions> = {
  1: { copyOnly: true },
  2: { series: true },
  6: { dateUnit: 'weekdays' },
  7: { dateUnit: 'months' },
  8: { dateUnit: 'years' },
  10: { growth: true },
};

const ISO = /^\d{4}-\d{2}-\d{2}$/;
const serialOf = (iso: string): number => Date.parse(`${iso}T00:00:00Z`) / 86_400_000 + 25_569;
const isoOf = (serial: number): string =>
  new Date(Math.round((serial - 25_569) * 86_400_000)).toISOString().slice(0, 10);

describe('series autofill matches desktop Excel', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale('ja-JP', 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  let slot = 0;
  const run = (
    grid: Seed[][],
    format: string | null,
    src: { r0: number; c0: number },
    dest: { r0: number; c0: number; r1: number; c1: number },
    options: FillOptions,
  ): Out[][] => {
    const base = slot++ * 12;
    const store = createSpreadsheetStore();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      const formats = new Map(s.format.formats);
      grid.forEach((line, r) => {
        line.forEach((seed, c) => {
          if (seed === null) return;
          const addr = { sheet: 0, row: base + src.r0 + r, col: src.c0 + c };
          if (typeof seed === 'number') {
            wb.setNumber(addr, seed);
            cells.set(addrKey(addr), { value: { kind: 'number', value: seed }, formula: null });
            if (format) formats.set(addrKey(addr), { numFmt: { kind: 'date', pattern: format } });
          } else if (ISO.test(seed)) {
            const value = serialOf(seed);
            wb.setNumber(addr, value);
            cells.set(addrKey(addr), { value: { kind: 'number', value }, formula: null });
            formats.set(addrKey(addr), { numFmt: { kind: 'date', pattern: format ?? 'yyyy/m/d' } });
          } else {
            wb.setText(addr, seed);
            cells.set(addrKey(addr), { value: { kind: 'text', value: seed }, formula: null });
          }
        });
      });
      return { ...s, data: { ...s.data, cells }, format: { ...s.format, formats } };
    });
    const srcRange: Range = {
      sheet: 0,
      r0: base + src.r0,
      c0: src.c0,
      r1: base + src.r0 + grid.length - 1,
      c1: src.c0 + (grid[0]?.length ?? 1) - 1,
    };
    const destRange: Range = {
      sheet: 0,
      r0: base + dest.r0,
      c0: dest.c0,
      r1: base + dest.r1,
      c1: dest.c1,
    };
    fillRange(store.getState(), wb, srcRange, destRange, { ...options, formatting: 'with', store });
    const out: Out[][] = [];
    for (let r = dest.r0; r <= dest.r1; r += 1) {
      const line: Out[] = [];
      for (let c = dest.c0; c <= dest.c1; c += 1) {
        const addr = { sheet: 0, row: base + r, col: c };
        const v = wb.getValue(addr);
        const dated = store.getState().format.formats.get(addrKey(addr))?.numFmt?.kind === 'date';
        line.push(
          v.kind === 'number'
            ? dated
              ? isoOf(v.value)
              : v.value
            : v.kind === 'text'
              ? v.value
              : null,
        );
      }
      out.push(line);
    }
    return out;
  };

  const sameOut = (got: Out[], want: Out[]): void => {
    expect(got).toHaveLength(want.length);
    want.forEach((expected, i) => {
      const actual = got[i];
      if (typeof expected === 'number' && typeof actual === 'number') {
        expect(
          Math.abs(actual - expected),
          `cell ${i}: ${actual} vs ${expected}`,
        ).toBeLessThanOrEqual(1e-5 * Math.max(1, Math.abs(expected)));
      } else {
        expect(actual, `cell ${i}`).toEqual(expected);
      }
    });
  };

  describe('numeric seeds down (linear trends)', () => {
    for (const c of fixture.linear) {
      it(`${JSON.stringify(c.seeds)} type ${c.type}`, () => {
        const n = c.seeds.length;
        const out = run(
          c.seeds.map((s) => [s]),
          null,
          { r0: 0, c0: 0 },
          { r0: 0, c0: 0, r1: n + 3, c1: 0 },
          OPTIONS[c.type] ?? {},
        );
        sameOut(
          out.map((row) => row[0] ?? null),
          c.values,
        );
      });
    }
  });

  describe('fills above and to the left of the seeds', () => {
    for (const c of fixture.reverse) {
      it(c.name, () => {
        const n = c.seeds.length;
        const options = OPTIONS[c.type ?? -1] ?? {};
        const seeds = c.seeds;
        let out: Out[];
        if (c.dir === 'up') {
          out = run(
            seeds.map((s) => [s]),
            c.format,
            { r0: 7 - n, c0: 0 },
            { r0: 0, c0: 0, r1: 6, c1: 0 },
            options,
          ).map((row) => row[0] ?? null);
        } else if (c.dir === 'left') {
          out =
            run(
              [seeds],
              c.format,
              { r0: 0, c0: 7 - n },
              { r0: 0, c0: 0, r1: 0, c1: 6 },
              options,
            )[0] ?? [];
        } else {
          out =
            run([seeds], c.format, { r0: 0, c0: 0 }, { r0: 0, c0: 0, r1: 0, c1: 6 }, options)[0] ??
            [];
        }
        sameOut(out, c.values as Out[]);
      });
    }
  });

  describe('multi-column and multi-row seed blocks', () => {
    for (const c of fixture.blocks) {
      it(c.name, () => {
        const h = c.grid.length;
        const w = c.grid[0]?.length ?? 1;
        const want = c.values as Out[][];
        const rows = want.length;
        const cols = want[0]?.length ?? 1;
        const out =
          c.dir === 'down'
            ? run(
                c.grid,
                null,
                { r0: 0, c0: 0 },
                { r0: 0, c0: 0, r1: rows - 1, c1: w - 1 },
                OPTIONS[c.type ?? -1] ?? {},
              )
            : run(
                c.grid,
                null,
                { r0: 0, c0: 0 },
                { r0: 0, c0: 0, r1: h - 1, c1: cols - 1 },
                OPTIONS[c.type ?? -1] ?? {},
              );
        want.forEach((row, r) => {
          sameOut(out[r] ?? [], Array.isArray(row) ? row : [row]);
        });
      });
    }
  });

  describe('growth trend fits', () => {
    const cases: Array<[number[], number[]]> = [
      [
        [2, 4],
        [8, 16, 32, 64],
      ],
      [
        [1, 2, 3],
        [5.451362, 9.442036, 16.354085, 28.326107],
      ],
      [
        [1, 3, 2],
        [3.634241, 5.139593, 7.268482, 10.279186],
      ],
      [
        [3, 2],
        [1.333333, 0.888889, 0.592593, 0.395062],
      ],
      [
        [-2, -4],
        [0, 0, 0, 0],
      ],
      [
        [0, 1, 2],
        [0, 0, 0, 0],
      ],
    ];
    for (const [seeds, want] of cases) {
      it(JSON.stringify(seeds), () => {
        const n = seeds.length;
        const out = run(
          seeds.map((s) => [s]),
          null,
          { r0: 0, c0: 0 },
          { r0: 0, c0: 0, r1: n + 3, c1: 0 },
          { growth: true },
        );
        sameOut(
          out.slice(n).map((row) => row[0] ?? null),
          want,
        );
      });
    }
  });
});
