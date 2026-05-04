import { describe, expect, it } from 'vitest';
import {
  defaultTableOverlay,
  isBandedRow,
  isHeaderRow,
  isTotalRow,
  removeTable,
  type TableOverlay,
  tableForCell,
  upsertTable,
} from '../../../src/commands/format-as-table.js';

const range = (r0: number, c0: number, r1: number, c1: number) =>
  ({ sheet: 0, r0, c0, r1, c1 }) as const;

describe('defaultTableOverlay', () => {
  it('produces sensible defaults (header on, banded on, total off)', () => {
    const t = defaultTableOverlay('tbl1', range(0, 0, 5, 3));
    expect(t.style).toBe('medium');
    expect(t.showHeader).toBe(true);
    expect(t.showTotal).toBe(false);
    expect(t.banded).toBe(true);
  });
});

describe('row classification', () => {
  const t: TableOverlay = {
    id: 'tbl1',
    range: range(0, 0, 5, 3),
    style: 'medium',
    showHeader: true,
    showTotal: true,
    banded: true,
  };

  it('isHeaderRow only true for the top row inside the rect', () => {
    expect(isHeaderRow(t, 0, 0)).toBe(true);
    expect(isHeaderRow(t, 0, 3)).toBe(true);
    expect(isHeaderRow(t, 1, 0)).toBe(false);
    expect(isHeaderRow(t, 0, 5)).toBe(false); // outside col range
  });

  it('isHeaderRow false when showHeader is off', () => {
    const off = { ...t, showHeader: false };
    expect(isHeaderRow(off, 0, 0)).toBe(false);
  });

  it('isTotalRow only true for the bottom row inside the rect', () => {
    expect(isTotalRow(t, 5, 0)).toBe(true);
    expect(isTotalRow(t, 4, 0)).toBe(false);
  });

  it('isBandedRow alternates between data rows', () => {
    // r0=0 (header), r1=5 (total). Data rows: 1, 2, 3, 4.
    // Banded = every other from data start (row 1 → "even" no, row 2 → "odd" yes).
    expect(isBandedRow(t, 1, 0)).toBe(false);
    expect(isBandedRow(t, 2, 0)).toBe(true);
    expect(isBandedRow(t, 3, 0)).toBe(false);
    expect(isBandedRow(t, 4, 0)).toBe(true);
  });

  it('isBandedRow excludes header + total rows', () => {
    expect(isBandedRow(t, 0, 0)).toBe(false);
    expect(isBandedRow(t, 5, 0)).toBe(false);
  });

  it('isBandedRow false when banded is off', () => {
    expect(isBandedRow({ ...t, banded: false }, 2, 0)).toBe(false);
  });
});

describe('tableForCell', () => {
  const t1: TableOverlay = {
    id: 'tbl1',
    range: range(0, 0, 5, 3),
    style: 'medium',
    showHeader: true,
    showTotal: false,
    banded: true,
  };
  const t2: TableOverlay = {
    id: 'tbl2',
    range: { sheet: 1, r0: 0, c0: 0, r1: 1, c1: 1 },
    style: 'light',
    showHeader: true,
    showTotal: false,
    banded: true,
  };

  it('returns the matching overlay for a cell inside the range', () => {
    expect(tableForCell([t1, t2], 0, 2, 1)?.id).toBe('tbl1');
  });

  it('returns null when sheet mismatches', () => {
    expect(tableForCell([t1], 1, 2, 1)).toBeNull();
  });

  it('returns null when cell sits outside both ranges', () => {
    expect(tableForCell([t1, t2], 0, 99, 99)).toBeNull();
  });

  it('returns the first overlay in registration order on overlap', () => {
    const overlap: TableOverlay = { ...t1, id: 'tbl-overlap' };
    expect(tableForCell([t1, overlap], 0, 1, 1)?.id).toBe('tbl1');
  });
});

describe('upsertTable / removeTable', () => {
  const a: TableOverlay = {
    id: 'a',
    range: range(0, 0, 1, 1),
    style: 'light',
    showHeader: true,
    showTotal: false,
    banded: true,
  };

  it('upsert adds a fresh overlay', () => {
    expect(upsertTable([], a)).toEqual([a]);
  });

  it('upsert replaces an existing overlay with the same id', () => {
    const updated = { ...a, banded: false };
    const next = upsertTable([a], updated);
    expect(next).toHaveLength(1);
    expect(next[0]?.banded).toBe(false);
  });

  it('removeTable returns the same reference when id does not match (cheap signal)', () => {
    const arr = [a];
    expect(removeTable(arr, 'missing')).toBe(arr);
  });

  it('removeTable strips the matching id', () => {
    expect(removeTable([a], 'a')).toEqual([]);
  });
});
