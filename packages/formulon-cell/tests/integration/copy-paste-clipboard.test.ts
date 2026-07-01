import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { copy } from '../../src/commands/clipboard/copy.js';
import { pasteTSV } from '../../src/commands/clipboard/paste.js';
import { pasteSpecial } from '../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../src/commands/clipboard/snapshot.js';
import { addrKey } from '../../src/engine/workbook-handle.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Mount the full stack and walk a copy → paste round-trip through the store
 * + workbook glue. This pins the regression surface around `commands/clipboard`,
 * `store/mutators.replaceCells`, and the protection gate inside `pasteTSV`.
 */
describe('integration: clipboard copy → paste', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => {
    sheet.dispose();
  });

  function syncToStore(): void {
    const { instance, workbook } = sheet;
    mutators.replaceCells(instance.store, workbook.cells(0));
  }

  it('copies a range to TSV and pastes it at a new origin', () => {
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 3);
    workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 4);
    syncToStore();

    // Select the 2×2 range.
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    const cp = copy(instance.store.getState());
    expect(cp).not.toBeNull();
    expect(cp?.tsv).toBe('1\t2\r\n3\t4');

    // Paste at D5.
    mutators.setActive(instance.store, { sheet: 0, row: 4, col: 3 });
    const res = pasteTSV(instance.store.getState(), workbook, cp?.tsv ?? '');
    expect(res).not.toBeNull();
    expect(workbook.getValue({ sheet: 0, row: 4, col: 3 })).toEqual({ kind: 'number', value: 1 });
    expect(workbook.getValue({ sheet: 0, row: 5, col: 4 })).toEqual({ kind: 'number', value: 4 });
  });

  it('paste auto-coerces a leading "=" string back into a formula', () => {
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
    workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 20);
    syncToStore();

    mutators.setActive(instance.store, { sheet: 0, row: 2, col: 0 });
    pasteTSV(instance.store.getState(), workbook, '=A1+A2');

    expect(workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=A1+A2');
  });

  it('refuses to paste empty payload', () => {
    const { instance, workbook } = sheet;
    const res = pasteTSV(instance.store.getState(), workbook, '');
    expect(res).toBeNull();
  });

  it('a 1M-cell range is refused (OOM guard)', () => {
    const { instance } = sheet;
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1_000_000, c1: 1 });
    expect(copy(instance.store.getState())).toBeNull();
  });

  it('copying a formula cell yields its formula text, not the computed value', () => {
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 6);
    workbook.setFormula({ sheet: 0, row: 1, col: 0 }, '=A1*2');
    workbook.recalc();
    syncToStore();

    mutators.setRange(instance.store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    const cp = copy(instance.store.getState());
    // The formula travels verbatim so a paste into Excel / Sheets rebuilds it.
    expect(cp?.tsv).toBe('=A1*2');
    // Suppress unused-var lint in addrKey import.
    expect(addrKey({ sheet: 0, row: 1, col: 0 })).toBe('0:1:0');
  });

  it('copy re-anchors relative refs by the paste offset', () => {
    const { instance, workbook } = sheet;
    workbook.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
    workbook.recalc();
    syncToStore();
    const snap = captureSnapshot(instance.store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 2,
      r1: 0,
      c1: 2,
    });
    expect(snap?.mode).toBe('copy');
    // Paste at C5 (row 4): copy shifts refs down by 4 rows.
    mutators.setActive(instance.store, { sheet: 0, row: 4, col: 2 });
    // biome-ignore lint/style/noNonNullAssertion: snapshot asserted above
    pasteSpecial(instance.store.getState(), instance.store, workbook, snap!, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    expect(workbook.cellFormula({ sheet: 0, row: 4, col: 2 })).toBe('=A5+B5');
  });

  it('cut pastes formulas verbatim — no re-relativization (C-4)', () => {
    const { instance, workbook } = sheet;
    workbook.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
    workbook.recalc();
    syncToStore();
    // Snapshot captured with cut mode (as the cut handler does).
    const snap = captureSnapshot(
      instance.store.getState(),
      { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
      'cut',
    );
    expect(snap?.mode).toBe('cut');
    mutators.setActive(instance.store, { sheet: 0, row: 4, col: 2 });
    // biome-ignore lint/style/noNonNullAssertion: snapshot asserted above
    pasteSpecial(instance.store.getState(), instance.store, workbook, snap!, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    // Excel keeps the formula intact on a move.
    expect(workbook.cellFormula({ sheet: 0, row: 4, col: 2 })).toBe('=A1+B1');
  });
});
