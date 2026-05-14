import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: mount the full Spreadsheet against the stub engine and walk
 * through edit → undo → redo. Verifies that the shared History attached during
 * mount drives both the engine state and the store-side cells projection in
 * lockstep — the regression surface that the recent mount/ refactor moved.
 */
describe('integration: edit / undo / redo', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => {
    sheet.dispose();
  });

  it('setNumber pushes onto undo, undo restores blank, redo re-applies', () => {
    const { instance, workbook } = sheet;
    const a1 = { sheet: 0, row: 0, col: 0 };

    workbook.setNumber(a1, 42);
    expect(workbook.getValue(a1)).toEqual({ kind: 'number', value: 42 });

    const undid = instance.undo();
    expect(undid).toBe(true);
    expect(workbook.getValue(a1)).toEqual({ kind: 'blank' });

    const redid = instance.redo();
    expect(redid).toBe(true);
    expect(workbook.getValue(a1)).toEqual({ kind: 'number', value: 42 });
  });

  it('undo returns false when there is nothing to undo', () => {
    expect(sheet.instance.undo()).toBe(false);
  });

  it('a setFormula → undo round-trip drops the formula on the engine', () => {
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    workbook.setFormula({ sheet: 0, row: 2, col: 0 }, '=A1+A2');

    expect(workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=A1+A2');

    instance.undo();
    expect(workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBeNull();
  });
});
