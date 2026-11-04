import { type FillOptions, fillRange } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import { formatCodeToNumFmt } from '../../../src/engine/format-writeback.js';
import type { Range } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { formatNumber } from '../../../src/format/number-format.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

export type Dir = 'down' | 'right' | 'up' | 'left';

export interface SeedCell {
  n?: number;
  t?: string;
  f?: string;
  b?: boolean;
}

/** One measured fill: the seed block, and the cells Excel produced beside it. */
export interface MatrixCase {
  id: string;
  dir: Dir;
  nf: string;
  cellNf?: string[];
  seedPos: number[];
  seed: (SeedCell | null)[];
  filled: [number, number | string | boolean | null, string, string][];
}

export interface MatrixFile<C extends MatrixCase> {
  helperIndex: number;
  cases: C[];
}

/** Lane (column for vertical fills, row for horizontal) of the n-th case of a direction. */
export class LaneAllocator {
  private readonly used: Record<Dir, number> = { down: 0, right: 0, up: 0, left: 0 };

  next(dir: Dir): number {
    const back = dir === 'up' || dir === 'left';
    return this.used[dir]++ + (back ? 100 : 0);
  }
}

const HORIZONTAL_BASE = 100;

/**
 * Seed a case on the workbook, run the fill and return one line per cell that
 * differs from the measured value, number format or (for numbers) display text.
 */
export function runMatrixCase(
  wb: WorkbookHandle,
  c: MatrixCase,
  lane: number,
  helperIndex: number,
  options: FillOptions,
  locale: string,
): string[] {
  const helper = helperIndex - 1;
  const vertical = c.dir === 'down' || c.dir === 'up';
  const laneIndex = vertical ? lane : HORIZONTAL_BASE + lane;
  const at = (p: number) =>
    vertical ? { sheet: 0, row: p - 1, col: lane } : { sheet: 0, row: laneIndex, col: p - 1 };
  const helperAt = (p: number) =>
    vertical ? { sheet: 0, row: p - 1, col: helper } : { sheet: 0, row: helper, col: p - 1 };

  const store = createSpreadsheetStore();
  const cells = new Map(store.getState().data.cells);
  const formats = new Map(store.getState().format.formats);
  const filledPos = c.filled.map(([p]) => p);
  const lo = Math.min(...filledPos, ...c.seedPos);
  const hi = Math.max(...filledPos, ...c.seedPos);
  if (c.seed.some((s) => s?.f)) {
    for (let p = 1; p <= hi; p += 1) wb.setNumber(helperAt(p), p * 10);
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
    } else if (s.b !== undefined) {
      wb.setBool(addr, s.b);
      cells.set(addrKey(addr), { value: { kind: 'bool', value: s.b }, formula: null });
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
  const rng = (x: { row: number; col: number }, y: { row: number; col: number }): Range => ({
    sheet: 0,
    r0: x.row,
    c0: x.col,
    r1: y.row,
    c1: y.col,
  });
  fillRange(store.getState(), wb, rng(at(first), at(last)), rng(at(lo), at(hi)), {
    ...options,
    formatting: 'with',
    store,
  });

  const diffs: string[] = [];
  for (const [p, ev, enf, et] of c.filled) {
    const addr = at(p);
    const v = wb.getValue(addr);
    const gotFmt = store.getState().format.formats.get(addrKey(addr))?.numFmt;
    let gv: number | string | boolean | null = null;
    let text = '';
    if (v.kind === 'number') {
      gv = v.value;
      text = formatNumber(v.value, gotFmt, locale);
    } else if (v.kind === 'text' || v.kind === 'bool') {
      gv = v.value;
      text = String(v.value);
    }
    const wantFmt = formatCodeToNumFmt(enf);
    const okV =
      typeof ev === 'number' && typeof gv === 'number'
        ? Math.abs(ev - gv) <= 1e-9 * Math.max(1, Math.abs(ev))
        : (ev ?? null) === gv || (ev === '' && gv === null);
    const okF = JSON.stringify(gotFmt ?? null) === JSON.stringify(wantFmt ?? null);
    // General-format text depends on column width; text and boolean cells render through the
    // format's text section, which this comparison leaves to the value check.
    const okT = enf === 'General' || typeof gv !== 'number' || text === et;
    if (!okV || !okF || !okT)
      diffs.push(
        `@${p}: want v=${ev} nf=${enf} t=${JSON.stringify(et)} | got v=${gv} nf=${JSON.stringify(gotFmt ?? null)} t=${JSON.stringify(text)}`,
      );
  }
  return diffs;
}
