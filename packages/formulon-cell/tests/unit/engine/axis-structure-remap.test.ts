import { describe, expect, it } from 'vitest';
import { addrKey, MAX_COL } from '../../../src/engine/address.js';
import {
  mapAxisAddress,
  remapAxisFormatOverride,
  remapAxisSnapshot,
  structuralKeySet,
} from '../../../src/engine/axis-structure-remap.js';
import type { SheetAxisFormatOverride, SheetFontAxisSnapshot } from '../../../src/engine/types.js';

const snapshot = (patch: Partial<SheetFontAxisSnapshot> = {}): SheetFontAxisSnapshot => ({
  columns: patch.columns ?? [],
  rows: patch.rows ?? [],
});

describe('axis structure remap', () => {
  it('moves sparse rows and leaves compact columns untouched', () => {
    const before = snapshot({
      columns: [{ first: 0, last: MAX_COL, styleXf: 4 }],
      rows: [
        { row: 2, styleXf: 8 },
        { row: 5, styleXf: 9 },
      ],
    });
    expect(remapAxisSnapshot(before, { axis: 'row', split: 2, delta: 1 })).toEqual(
      snapshot({
        columns: [{ first: 0, last: MAX_COL, styleXf: 4 }],
        rows: [
          { row: 3, styleXf: 8 },
          { row: 6, styleXf: 9 },
        ],
      }),
    );
    expect(remapAxisSnapshot(before, { axis: 'row', split: 2, delta: -1 })).toEqual(
      snapshot({
        columns: [{ first: 0, last: MAX_COL, styleXf: 4 }],
        rows: [{ row: 4, styleXf: 9 }],
      }),
    );
  });

  it('clips full-width and partial column spans without enumerating columns', () => {
    const before = snapshot({
      columns: [
        { first: 0, last: MAX_COL, styleXf: 4 },
        { first: 2, last: 5, styleXf: 7 },
      ],
    });
    expect(remapAxisSnapshot(before, { axis: 'col', split: 2, delta: 1 }).columns).toEqual([
      { first: 0, last: MAX_COL, styleXf: 4 },
      { first: 3, last: 6, styleXf: 7 },
    ]);
    expect(
      remapAxisSnapshot(snapshot({ columns: [{ first: 0, last: 5, styleXf: 7 }] }), {
        axis: 'col',
        split: 2,
        delta: -2,
      }).columns,
    ).toEqual([{ first: 0, last: 3, styleXf: 7 }]);
  });

  it('drops deleted axis entries and direct keys while retaining undo metadata', () => {
    const override: SheetAxisFormatOverride = {
      baseline: snapshot({ rows: [{ row: 2, styleXf: 3 }] }),
      desired: snapshot({ rows: [{ row: 2, styleXf: 8 }] }),
      directXfs: [
        {
          key: '0:2:3',
          before: { directXfIndex: 0, effectiveXfIndex: 3, source: 1 },
          afterXfIndex: 9,
        },
      ],
    };
    const mapped = remapAxisFormatOverride(override, { axis: 'row', split: 2, delta: 1 });
    expect(mapped.desired.rows).toEqual([{ row: 3, styleXf: 8 }]);
    expect(mapped.directXfs?.[0]?.key).toBe('0:3:3');
    expect(mapped.directXfs?.[0]?.before).toEqual(override.directXfs?.[0]?.before);
    expect(
      remapAxisFormatOverride(override, { axis: 'row', split: 2, delta: -1 }).directXfs,
    ).toEqual([]);
  });

  it('returns bounded before and mapped-after checkpoint keys', () => {
    const shift = { axis: 'col' as const, split: 2, delta: 1 };
    const keys = structuralKeySet(['0:1:1', '0:1:3', '0:2:4'], 0, shift);
    expect(keys).toEqual(new Set(['0:1:1', '0:1:3', '0:2:4', '0:1:4', '0:2:5']));
    expect(mapAxisAddress({ sheet: 0, row: 2, col: 2 }, shift)).toEqual({
      sheet: 0,
      row: 2,
      col: 3,
    });
    expect(addrKey({ sheet: 0, row: 2, col: 3 })).toBe('0:2:3');
  });
});
