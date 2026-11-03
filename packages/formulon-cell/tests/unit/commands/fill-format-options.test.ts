import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { type FillOptions, fillRange } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { CellFormat } from '../../../src/store/types.js';

// Desktop Excel (ja), Range.AutoFill with xlFillDefault / Copy / Series /
// Formats / Values from styled seeds; `Formats` is Fill Formatting Only and
// `Values` is Fill Without Formatting.
const STYLE: CellFormat = {
  numFmt: { kind: 'fixed', decimals: 2 },
  bold: true,
  fill: '#ffc800',
  color: '#c80000',
};
const BOLD_FILL: CellFormat = { bold: true, fill: '#00ff00' };

type Kind = 'default' | 'copy' | 'series' | 'formats' | 'values';
const OPTIONS: Record<Kind, FillOptions> = {
  default: {},
  copy: { copyOnly: true },
  series: { series: true },
  formats: { formatting: 'only' },
  values: { formatting: 'without' },
};

type Cell = number | string | null;
const styled = (f: CellFormat | undefined, style: CellFormat): boolean =>
  f?.bold === style.bold && f?.fill === style.fill && f?.color === style.color;
const plain = (f: CellFormat | undefined): boolean =>
  !f?.bold && !f?.fill && !f?.color && !f?.numFmt;

describe('autofill formatting options match desktop Excel', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale('ja-JP', 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  let slot = 0;
  interface Setup {
    seeds: Array<{ row: number; value: number | string | { f: string }; format: CellFormat }>;
    existing?: Array<{ row: number; value: string }>;
    src: [number, number];
    dest: [number, number];
  }
  const run = (setup: Setup, kind: Kind) => {
    const col = slot++;
    const store = createSpreadsheetStore();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      const formats = new Map(s.format.formats);
      for (const seed of setup.seeds) {
        const addr = { sheet: 0, row: seed.row, col };
        if (typeof seed.value === 'number') {
          wb.setNumber(addr, seed.value);
          cells.set(addrKey(addr), { value: { kind: 'number', value: seed.value }, formula: null });
        } else if (typeof seed.value === 'string') {
          wb.setText(addr, seed.value);
          cells.set(addrKey(addr), { value: { kind: 'text', value: seed.value }, formula: null });
        } else {
          wb.setFormula(addr, seed.value.f);
          cells.set(addrKey(addr), { value: { kind: 'number', value: 0 }, formula: seed.value.f });
        }
        formats.set(addrKey(addr), seed.format);
      }
      for (const e of setup.existing ?? []) {
        const addr = { sheet: 0, row: e.row, col };
        wb.setText(addr, e.value);
        cells.set(addrKey(addr), { value: { kind: 'text', value: e.value }, formula: null });
      }
      return { ...s, data: { ...s.data, cells }, format: { ...s.format, formats } };
    });
    wb.recalc();
    const range = (r: [number, number]): Range => ({
      sheet: 0,
      r0: r[0],
      c0: col,
      r1: r[1],
      c1: col,
    });
    fillRange(store.getState(), wb, range(setup.src), range(setup.dest), {
      ...OPTIONS[kind],
      store,
    });
    wb.recalc();
    const at = (row: number) => {
      const addr = { sheet: 0, row, col };
      const v = wb.getValue(addr);
      const value: Cell = v.kind === 'number' || v.kind === 'text' ? v.value : null;
      return {
        value,
        formula: wb.cellFormula(addr),
        format: store.getState().format.formats.get(addrKey(addr)),
      };
    };
    return at;
  };

  const pair: Setup = {
    seeds: [
      { row: 0, value: 1, format: STYLE },
      { row: 1, value: 2, format: STYLE },
    ],
    src: [0, 1],
    dest: [0, 5],
  };

  it.each([
    ['default', [1, 2, 3, 4, 6]],
    ['series', [1, 2, 3, 4, 6]],
    ['values', [1, 2, 3, 4, 6]],
  ] as [Kind, number[]][])('number pair, %s: series values', (kind, values) => {
    const at = run(pair, kind);
    expect([0, 1, 2, 3, 5].map((r) => at(r).value)).toEqual(values);
  });

  it.each(['default', 'copy', 'series', 'formats'] as Kind[])(
    'number pair, %s: filled cells carry the seed format',
    (kind) => {
      const at = run(pair, kind);
      for (const r of [2, 3, 5]) expect(styled(at(r).format, STYLE), `row ${r}`).toBe(true);
      expect(at(2).format?.numFmt).toEqual(STYLE.numFmt);
    },
  );

  it('number pair, copy: the pair repeats', () => {
    const at = run(pair, 'copy');
    expect([2, 3, 4, 5].map((r) => at(r).value)).toEqual([1, 2, 1, 2]);
  });

  it('number pair, formats only: values stay as they were, formats applied', () => {
    const at = run(pair, 'formats');
    expect([2, 3, 4, 5].map((r) => at(r).value)).toEqual([null, null, null, null]);
  });

  it('number pair, values only: the filled cells are unformatted', () => {
    const at = run(pair, 'values');
    for (const r of [2, 3, 5]) expect(plain(at(r).format), `row ${r}`).toBe(true);
  });

  it('formats only leaves existing destination text alone', () => {
    const at = run(
      {
        seeds: [{ row: 0, value: 5, format: STYLE }],
        existing: [
          { row: 1, value: 'x' },
          { row: 2, value: 'y' },
          { row: 3, value: 'z' },
        ],
        src: [0, 0],
        dest: [0, 3],
      },
      'formats',
    );
    expect([1, 2, 3].map((r) => at(r).value)).toEqual(['x', 'y', 'z']);
    for (const r of [1, 2, 3]) expect(styled(at(r).format, STYLE)).toBe(true);
  });

  it('a single number seed copies under values only, unformatted', () => {
    const at = run(
      { seeds: [{ row: 0, value: 5, format: STYLE }], src: [0, 0], dest: [0, 3] },
      'values',
    );
    expect([1, 2, 3].map((r) => at(r).value)).toEqual([5, 5, 5]);
    for (const r of [1, 2, 3]) expect(plain(at(r).format)).toBe(true);
  });

  describe('formula seed', () => {
    const formula = (kind: Kind) =>
      run(
        {
          seeds: [
            {
              row: 0,
              value: { f: '=A1*2' },
              format: { bold: true, numFmt: { kind: 'fixed', decimals: 0, thousands: true } },
            },
          ],
          src: [0, 0],
          dest: [0, 3],
        },
        kind,
      );
    it.each(['default', 'copy', 'series', 'values'] as Kind[])('%s shifts the formula', (kind) => {
      const at = formula(kind);
      expect(at(1).formula?.replace(/[A-Z]+/g, 'X')).toBe('=X2*2');
      expect(at(3).formula?.replace(/[A-Z]+/g, 'X')).toBe('=X4*2');
    });
    it('formats only writes no formula', () => {
      const at = formula('formats');
      expect(at(1).formula ?? null).toBeNull();
      expect(at(1).format?.bold).toBe(true);
    });
    it('values only fills formulas without the format', () => {
      const at = formula('values');
      expect(at(1).format?.bold ?? false).toBe(false);
    });
  });

  describe('text seed', () => {
    const text = (kind: Kind) =>
      run(
        { seeds: [{ row: 0, value: 'Item1', format: BOLD_FILL }], src: [0, 0], dest: [0, 2] },
        kind,
      );
    it('default and series continue the numbered text with the format', () => {
      for (const kind of ['default', 'series'] as Kind[]) {
        const at = text(kind);
        expect([1, 2].map((r) => at(r).value)).toEqual(['Item2', 'Item3']);
        expect(styled(at(1).format, BOLD_FILL)).toBe(true);
      }
    });
    it('copy repeats the text with the format', () => {
      const at = text('copy');
      expect([1, 2].map((r) => at(r).value)).toEqual(['Item1', 'Item1']);
      expect(styled(at(2).format, BOLD_FILL)).toBe(true);
    });
    it('formats only leaves the values empty', () => {
      const at = text('formats');
      expect([1, 2].map((r) => at(r).value)).toEqual([null, null]);
      expect(styled(at(1).format, BOLD_FILL)).toBe(true);
    });
    it('values only continues the series unformatted', () => {
      const at = text('values');
      expect([1, 2].map((r) => at(r).value)).toEqual(['Item2', 'Item3']);
      expect(plain(at(1).format)).toBe(true);
    });
  });

  it('upward fill from a single bold seed: formats only leaves values empty, values only copies unformatted', () => {
    const setup: Setup = {
      seeds: [{ row: 5, value: 7, format: { bold: true } }],
      src: [5, 5],
      dest: [2, 5],
    };
    const only = run(setup, 'formats');
    expect([2, 3, 4].map((r) => only(r).value)).toEqual([null, null, null]);
    expect(only(3).format?.bold).toBe(true);
    const values = run(setup, 'values');
    expect([2, 3, 4].map((r) => values(r).value)).toEqual([7, 7, 7]);
    expect(values(3).format?.bold ?? false).toBe(false);
  });
});
