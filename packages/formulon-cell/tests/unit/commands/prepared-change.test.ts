import { describe, expect, it } from 'vitest';
import {
  changedFormatAddresses,
  effectivePendingFormatAddresses,
  type PreparedChange,
  projectPreparedFormats,
  sameStructuredValue,
  unionAddresses,
} from '../../../src/commands/prepared-change.js';
import { addrKey } from '../../../src/engine/address.js';
import type { Addr } from '../../../src/engine/types.js';
import type { CellFormat } from '../../../src/store/types.js';

const a1: Addr = { sheet: 0, row: 0, col: 0 };
const b1: Addr = { sheet: 0, row: 0, col: 1 };

const change = (addr: Addr, extra: Partial<PreparedChange> = {}): PreparedChange => ({
  patch: { addr, value: { kind: 'blank' }, formula: null },
  ...extra,
});

const percent: CellFormat['numFmt'] = { kind: 'percent', decimals: 0 };

describe('sameStructuredValue', () => {
  it('compares nested objects and arrays by structure, ignoring key order', () => {
    expect(sameStructuredValue({ a: 1, b: [1, { c: 2 }] }, { b: [1, { c: 2 }], a: 1 })).toBe(true);
    expect(sameStructuredValue({ a: 1 }, { a: 2 })).toBe(false);
    expect(sameStructuredValue([1, 2], [1, 2, 3])).toBe(false);
    expect(sameStructuredValue([], {})).toBe(false);
    expect(sameStructuredValue(undefined, undefined)).toBe(true);
    expect(sameStructuredValue(null, {})).toBe(false);
  });
});

describe('projectPreparedFormats', () => {
  it('applies the implicit number format only over a general or missing format', () => {
    const before = new Map<string, CellFormat>([
      [addrKey(a1), { numFmt: { kind: 'general' } }],
      [addrKey(b1), { numFmt: { kind: 'fixed', decimals: 2 } }],
    ]);
    const out = projectPreparedFormats(before, [
      change(a1, { implicitFormat: percent }),
      change(b1, { implicitFormat: percent }),
      change({ sheet: 0, row: 5, col: 5 }, { implicitFormat: percent }),
    ]);
    expect(out.get(addrKey(a1))?.numFmt).toEqual(percent);
    expect(out.get(addrKey(b1))?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
    expect(out.get('0:5:5')).toEqual({ numFmt: percent });
  });

  it('lets a pending format override the implicit format and merges borders', () => {
    const before = new Map<string, CellFormat>([
      [addrKey(a1), { bold: true, borders: { top: { style: 'thin' } } }],
    ]);
    const out = projectPreparedFormats(before, [
      change(a1, {
        implicitFormat: percent,
        pendingFormat: { numFmt: { kind: 'general' }, borders: { bottom: { style: 'thin' } } },
      }),
    ]);
    const next = out.get(addrKey(a1));
    expect(next?.bold).toBe(true);
    expect(next?.numFmt).toEqual({ kind: 'general' });
    expect(next?.borders).toEqual({ top: { style: 'thin' }, bottom: { style: 'thin' } });
  });

  it('does not mutate the input map and leaves untouched cells absent', () => {
    const before = new Map<string, CellFormat>();
    const out = projectPreparedFormats(before, [change(a1)]);
    expect(before.size).toBe(0);
    expect(out.size).toBe(0);
  });
});

describe('effectivePendingFormatAddresses', () => {
  it('ignores implicit formats on their own', () => {
    const prepared = [change(a1, { implicitFormat: percent })];
    expect(effectivePendingFormatAddresses(new Map(), prepared)).toEqual([]);
  });

  it('reports a pending format that changes the projected result', () => {
    const prepared = [
      change(a1, { pendingFormat: { bold: true } }),
      change(b1, { implicitFormat: percent }),
    ];
    expect(effectivePendingFormatAddresses(new Map(), prepared)).toEqual([a1]);
  });

  it('drops a pending format that matches the implicit format already projected', () => {
    const prepared = [change(a1, { implicitFormat: percent, pendingFormat: { numFmt: percent } })];
    expect(effectivePendingFormatAddresses(new Map(), prepared)).toEqual([]);
  });

  it('drops a pending format that equals the existing format', () => {
    const before = new Map<string, CellFormat>([[addrKey(a1), { bold: true }]]);
    const prepared = [change(a1, { pendingFormat: { bold: true } })];
    expect(effectivePendingFormatAddresses(before, prepared)).toEqual([]);
  });
});

describe('changedFormatAddresses', () => {
  it('lists each changed address once, in prepared order', () => {
    const before = new Map<string, CellFormat>();
    const after = new Map<string, CellFormat>([
      [addrKey(a1), { bold: true }],
      [addrKey(b1), { italic: true }],
    ]);
    expect(changedFormatAddresses(before, after, [change(b1), change(a1), change(b1)])).toEqual([
      b1,
      a1,
    ]);
    expect(changedFormatAddresses(after, after, [change(a1)])).toEqual([]);
  });
});

describe('unionAddresses', () => {
  it('deduplicates by address key and keeps first-seen order', () => {
    expect(unionAddresses([a1, b1], [b1, { ...a1 }], [])).toEqual([a1, b1]);
  });
});
