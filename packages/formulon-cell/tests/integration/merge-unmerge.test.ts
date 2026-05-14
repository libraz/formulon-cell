import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { applyMerge, applyUnmerge, expandRangeWithMerges } from '../../src/commands/merge.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: merge / unmerge against the live mount. Covers the WASM-side
 * blanking of non-anchor cells, undo round-trip, and the expandRangeWithMerges
 * helper that selection-driven commands rely on.
 */
describe('integration: merge / unmerge', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('applyMerge blanks non-anchor cells and records a single merge', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'A1');
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'B1');
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'A2');
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));

    const ok = applyMerge(instance.store, workbook, instance.history, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    expect(ok).toBe(true);

    // Anchor keeps its value; the other 3 are blanked.
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'A1' });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });

    expect(instance.store.getState().merges.byAnchor.size).toBe(1);
  });

  it('applyMerge on a 1x1 range is a no-op', () => {
    const { instance, workbook } = sheet;
    expect(
      applyMerge(instance.store, workbook, instance.history, {
        sheet: 0,
        r0: 2,
        c0: 2,
        r1: 2,
        c1: 2,
      }),
    ).toBe(false);
    expect(instance.store.getState().merges.byAnchor.size).toBe(0);
  });

  it('undo restores the cells blanked by applyMerge', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'A1');
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'B1');
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));

    applyMerge(instance.store, workbook, instance.history, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 1,
    });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });

    expect(instance.undo()).toBe(true);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'B1' });
    expect(instance.store.getState().merges.byAnchor.size).toBe(0);
  });

  it('applyUnmerge clears an intersecting merge', () => {
    const { instance, workbook } = sheet;
    applyMerge(instance.store, workbook, instance.history, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    expect(instance.store.getState().merges.byAnchor.size).toBe(1);

    const ok = applyUnmerge(instance.store, workbook, instance.history, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    expect(ok).toBe(true);
    expect(instance.store.getState().merges.byAnchor.size).toBe(0);
  });

  it('expandRangeWithMerges grows a selection to enclose every overlapping merge', () => {
    const { instance, workbook } = sheet;
    applyMerge(instance.store, workbook, instance.history, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 2,
    });
    // A selection of just B1 should expand to the full 3x3 merge.
    const expanded = expandRangeWithMerges(instance.store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 1,
      r1: 0,
      c1: 1,
    });
    expect(expanded).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
  });
});
