import { describe, expect, it } from 'vitest';
import {
  computeNamedCellStyles,
  type NamedCellStylesView,
} from '../../../src/engine/cell-styles-meta.js';

type EngineStyle = NonNullable<ReturnType<NamedCellStylesView['getCellStyle']>>;

const fullStyle = (overrides: Partial<EngineStyle> = {}): EngineStyle => ({
  name: 'Normal',
  xfId: 0,
  builtinId: 0,
  iLevel: 0,
  hidden: false,
  customBuiltin: false,
  ...overrides,
});

const view = (styles: (EngineStyle | null)[]): NamedCellStylesView => ({
  cellStyleCount: () => styles.length,
  getCellStyle: (i: number) => styles[i] ?? null,
});

describe('computeNamedCellStyles', () => {
  it('returns an empty list when count is zero', () => {
    expect(computeNamedCellStyles(view([]))).toEqual([]);
  });

  it('emits one entry per non-hidden style with its index', () => {
    expect(
      computeNamedCellStyles(
        view([
          fullStyle({ name: 'Normal', xfId: 0, builtinId: 0 }),
          fullStyle({ name: 'Heading 1', xfId: 1, builtinId: 16, iLevel: 1 }),
        ]),
      ),
    ).toEqual([
      {
        index: 0,
        name: 'Normal',
        xfId: 0,
        builtinId: 0,
        iLevel: 0,
        customBuiltin: false,
      },
      {
        index: 1,
        name: 'Heading 1',
        xfId: 1,
        builtinId: 16,
        iLevel: 1,
        customBuiltin: false,
      },
    ]);
  });

  it('filters out hidden built-ins (the gallery hides those)', () => {
    const names = computeNamedCellStyles(
      view([
        fullStyle({ name: 'Normal' }),
        fullStyle({ name: 'Comma [0]', hidden: true }),
        fullStyle({ name: 'Heading 1' }),
      ]),
    ).map((s) => s.name);
    expect(names).toEqual(['Normal', 'Heading 1']);
  });

  it('skips slots where getCellStyle returns null (out-of-range guard)', () => {
    expect(computeNamedCellStyles(view([null, fullStyle({ name: 'Total' })]))).toHaveLength(1);
  });

  it('preserves the original index even when earlier slots are filtered', () => {
    const out = computeNamedCellStyles(
      view([
        fullStyle({ name: 'Normal' }),
        fullStyle({ name: 'Hidden', hidden: true }),
        fullStyle({ name: 'Heading 1', xfId: 5 }),
      ]),
    );
    expect(out.map((s) => s.index)).toEqual([0, 2]);
  });
});
