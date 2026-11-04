import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { type FillOptions, fillRange } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import { formatCodeToNumFmt } from '../../../src/engine/format-writeback.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { formatNumber } from '../../../src/format/number-format.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type Dir = 'down' | 'right' | 'up' | 'left';
interface SeedCell {
  n?: number;
  t?: string;
  f?: string;
}
interface Case {
  id: string;
  dir: Dir;
  typ: number | null;
  nf: string;
  cellNf?: string[];
  seedPos: number[];
  seed: (SeedCell | null)[];
  filled: [number, number | string | null, string, string][];
}

const matrix = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill-edge-matrix.json'), 'utf8'),
) as { helperIndex: number; cases: Case[] };

const LOCALE = 'ja-JP';
const HELPER = matrix.helperIndex - 1;
// A plain fill-handle drag equals the default fill type, so `null` carries no options.
const OPTIONS: Record<number, FillOptions> = {
  1: { copyOnly: true },
  2: { series: true },
  5: { dateUnit: 'days' },
  6: { dateUnit: 'weekdays' },
  7: { dateUnit: 'months' },
  8: { dateUnit: 'years' },
  10: { growth: true },
};

describe('fillRange matches the measured AutoFill edge matrix (ja-JP)', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale(LOCALE, 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  const lanes: Record<Dir, number> = { down: 0, right: 0, up: 0, left: 0 };
  const horizontalBase = 100;
  matrix.cases.forEach((c, index) => {
    const vertical = c.dir === 'down' || c.dir === 'up';
    const lane = lanes[c.dir]++ + (c.dir === 'up' ? 100 : 0) + (c.dir === 'left' ? 100 : 0);
    const laneIndex = vertical ? lane : horizontalBase + lane;
    const at = (p: number) =>
      vertical ? { sheet: 0, row: p - 1, col: lane } : { sheet: 0, row: laneIndex, col: p - 1 };
    const helperAt = (p: number) =>
      vertical ? { sheet: 0, row: p - 1, col: HELPER } : { sheet: 0, row: HELPER, col: p - 1 };
    const label = `${index} ${c.id} ${c.dir} typ=${c.typ}`;

    it(label, () => {
      if (c.dir === 'up' || c.dir === 'left') expect(lane).toBeGreaterThanOrEqual(100);
      const store = createSpreadsheetStore();
      const cells = new Map(store.getState().data.cells);
      const formats = new Map(store.getState().format.formats);
      const filledMax = Math.max(...c.filled.map(([p]) => p));
      const filledMin = Math.min(...c.filled.map(([p]) => p));
      const needsHelper = c.seed.some((s) => s?.f);
      if (needsHelper) {
        for (let p = filledMin; p <= filledMax; p += 1) wb.setNumber(helperAt(p), p * 10);
      }
      c.seed.forEach((s, i) => {
        const addr = at(c.seedPos[i] as number);
        const numFmt = formatCodeToNumFmt(c.cellNf?.[i] ?? c.nf);
        if (numFmt) formats.set(addrKey(addr), { numFmt });
        if (!s) return;
        if (s.f !== undefined) {
          wb.setFormula(addr, s.f);
          cells.set(addrKey(addr), { value: wb.getValue(addr), formula: s.f });
        } else if (s.t !== undefined) {
          wb.setText(addr, s.t);
          cells.set(addrKey(addr), { value: { kind: 'text', value: s.t }, formula: null });
        } else {
          wb.setNumber(addr, s.n as number);
          cells.set(addrKey(addr), {
            value: { kind: 'number', value: s.n as number },
            formula: null,
          });
        }
      });
      store.setState((st) => ({
        ...st,
        data: { ...st.data, cells },
        format: { ...st.format, formats },
      }));

      const first = c.seedPos[0] as number;
      const last = c.seedPos[c.seedPos.length - 1] as number;
      const lo = Math.min(first, filledMin);
      const hi = Math.max(last, filledMax);
      const a = at(lo);
      const b = at(hi);
      const rng = (x: typeof a, y: typeof a): Range => ({
        sheet: 0,
        r0: x.row,
        c0: x.col,
        r1: y.row,
        c1: y.col,
      });
      fillRange(store.getState(), wb, rng(at(first), at(last)), rng(a, b), {
        ...(c.typ === null ? {} : OPTIONS[c.typ]),
        formatting: 'with',
        store,
      });

      const diffs: string[] = [];
      for (const [p, ev, enf, et] of c.filled) {
        const addr = at(p);
        const v = wb.getValue(addr);
        const gotFmt = store.getState().format.formats.get(addrKey(addr))?.numFmt;
        let gv: number | string | null = null;
        let text = '';
        if (v.kind === 'number') {
          gv = v.value;
          text = formatNumber(v.value, gotFmt, LOCALE);
        } else if (v.kind === 'text') {
          gv = v.value;
          text = v.value;
        }
        const wantFmt = formatCodeToNumFmt(enf);
        const okV =
          typeof ev === 'number' && typeof gv === 'number'
            ? Math.abs(ev - gv) <= 1e-9 * Math.max(1, Math.abs(ev))
            : (ev ?? null) === gv || (ev === '' && gv === null);
        const okF = JSON.stringify(gotFmt ?? null) === JSON.stringify(wantFmt ?? null);
        // Display text is checked where the number formatter, not the fill, decides it: General
        // depends on column width, the built-in m/d/yy follows the locale, and a negative value
        // under a literal-prefixed code puts its sign in front of the literal.
        const okT = enf === 'General' || enf === 'm/d/yy' || et.startsWith('-') || text === et;
        if (!okV || !okF || !okT)
          diffs.push(
            `@${p}: want v=${ev} nf=${enf} t=${JSON.stringify(et)} | got v=${gv} nf=${JSON.stringify(gotFmt ?? null)} t=${JSON.stringify(text)}`,
          );
      }
      expect(diffs, diffs.join('\n')).toEqual([]);
    });
  });
});
