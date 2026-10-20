import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { coerceInput } from '../../../src/commands/coerce-input.js';
import { type FillOptions, fillRange } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import { formatCodeToNumFmt } from '../../../src/engine/format-writeback.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type Cell =
  | { kind: 'number'; value: number; format?: string }
  | { kind: 'text' | 'error'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'blank' };

interface Case {
  id: string;
  source: string[];
  dir: 'down' | 'up' | 'right' | 'left';
  count: number;
  type?: string;
}

interface Oracle {
  cases: Case[];
  locales: Record<string, Record<string, { source: Cell[]; fill: Cell[] }>>;
}

// Captured from desktop Excel by scripts/excel-autofill-probe.py.
const oracle = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill.json'), 'utf8'),
) as Oracle;

// Trend fills come from the Series dialog, which offers no trend types.
const UNSUPPORTED_TYPES = new Set(['linear_trend', 'growth_trend']);
// Outside CJK locales the desktop application steps a full-width digit by its
// code point (`Item１` becomes `Item65250`); it is copied here instead.
const NOT_REPLICATED: Record<string, readonly string[]> = {
  'en-US': ['w8'],
  'de-DE': ['w8'],
  'fr-FR': ['w8'],
  'th-TH': ['w8'],
};

const OPTIONS: Record<string, FillOptions> = {
  fill_copy: { copyOnly: true },
  fill_series: { series: true },
  fill_days: { dateUnit: 'days' },
  fill_weekdays: { dateUnit: 'weekdays' },
  fill_months: { dateUnit: 'months' },
  fill_years: { dateUnit: 'years' },
};

/** Source range and filled cells (in fill order), 0-based, one case per column or row. */
const layout = (c: Case, column: number, row: number) => {
  const k = c.source.length;
  const back = c.dir === 'up' || c.dir === 'left';
  const along = Array.from({ length: k }, (_, i) => (back ? c.count : 0) + i);
  const filled = Array.from({ length: c.count }, (_, i) => (back ? c.count - 1 - i : k + i));
  const vertical = c.dir === 'down' || c.dir === 'up';
  const at = (p: number) => (vertical ? { row: p, col: column } : { row, col: p });
  const src = along.map(at);
  const span = [0, k + c.count - 1].map(at);
  const range = (a: { row: number; col: number }, b: { row: number; col: number }): Range => ({
    sheet: 0,
    r0: a.row,
    c0: a.col,
    r1: b.row,
    c1: b.col,
  });
  return {
    src,
    filled: filled.map(at),
    srcRange: range(src[0] as never, src[k - 1] as never),
    destRange: range(span[0] as never, span[1] as never),
  };
};

describe('fillRange matches desktop Excel AutoFill', () => {
  for (const [locale, results] of Object.entries(oracle.locales)) {
    describe(locale, () => {
      let wb: WorkbookHandle;
      beforeAll(async () => {
        wb = await WorkbookHandle.createDefault();
        const profile = spreadsheetProfileForLocale(locale, 'mac');
        if (profile) wb.setSpreadsheetProfileId(profile);
      });
      afterAll(() => wb.dispose());

      let column = 0;
      let row = 60;
      for (const c of oracle.cases) {
        const vertical = c.dir === 'down' || c.dir === 'up';
        const place = layout(c, vertical ? column++ : 0, vertical ? 0 : row++);
        if (UNSUPPORTED_TYPES.has(c.type ?? '')) continue;
        if (NOT_REPLICATED[locale]?.includes(c.id)) continue;
        const want = results[c.id];
        if (!want) continue;
        it(`${c.id} ${JSON.stringify(c.source)} ${c.dir}${c.type ? ` ${c.type}` : ''}`, () => {
          const store = createSpreadsheetStore();
          store.setState((s) => {
            const cells = new Map(s.data.cells);
            const formats = new Map(s.format.formats);
            c.source.forEach((text, i) => {
              const addr = { sheet: 0, ...(place.src[i] as { row: number; col: number }) };
              const input = coerceInput(text, { locale });
              if (input.kind === 'number') {
                wb.setNumber(addr, input.value);
                cells.set(addrKey(addr), {
                  value: { kind: 'number', value: input.value },
                  formula: null,
                });
                if (input.implicitFormat)
                  formats.set(addrKey(addr), { numFmt: input.implicitFormat });
              } else if (input.kind === 'text') {
                wb.setText(addr, input.value);
                cells.set(addrKey(addr), {
                  value: { kind: 'text', value: input.value },
                  formula: null,
                });
              } else if (input.kind === 'bool') {
                wb.setBool(addr, input.value);
                cells.set(addrKey(addr), {
                  value: { kind: 'bool', value: input.value },
                  formula: null,
                });
              } else {
                throw new Error(`unexpected source kind ${input.kind}`);
              }
            });
            return { ...s, data: { ...s.data, cells }, format: { ...s.format, formats } };
          });
          fillRange(store.getState(), wb, place.srcRange, place.destRange, {
            ...OPTIONS[c.type ?? ''],
            formatting: 'with',
            store,
          });
          const got = place.filled.map((at) => {
            const addr = { sheet: 0, ...at };
            const value = wb.getValue(addr);
            const numFmt = store.getState().format.formats.get(addrKey(addr))?.numFmt;
            return { value, numFmt };
          });
          want.fill.forEach((expected, i) => {
            const actual = got[i];
            if (expected.kind === 'number') {
              expect(actual?.value.kind, `cell ${i}`).toBe('number');
              const value = actual?.value.kind === 'number' ? actual.value.value : Number.NaN;
              expect(Math.abs(value - expected.value), `cell ${i}: ${value}`).toBeLessThanOrEqual(
                1e-9 * Math.max(1, Math.abs(expected.value)),
              );
              expect(actual?.numFmt ?? null, `format ${i}`).toEqual(
                expected.format ? formatCodeToNumFmt(expected.format) : null,
              );
            } else if (expected.kind === 'blank') {
              expect(actual?.value.kind, `cell ${i}`).toBe('blank');
            } else {
              expect(actual?.value, `cell ${i}`).toEqual(expected);
            }
          });
        });
      }
    });
  }
});
