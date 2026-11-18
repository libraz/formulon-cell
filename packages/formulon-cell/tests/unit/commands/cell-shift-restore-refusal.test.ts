// @vitest-environment node

import { CfRuleType, FormulonError, StatusCode } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  captureShiftReplayCheckpoint,
  deleteCells,
  restoreShiftReplayCheckpoint,
} from '../../../src/commands/cell-shift.js';
import { History } from '../../../src/commands/history.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// A write the engine refuses partway through a restore must leave the sheet
// as the restore found it, never half rewritten.

const BAND: Range = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 }; // B2:C3

const live: WorkbookHandle[] = [];
afterEach(() => {
  vi.restoreAllMocks();
  for (const wb of live.splice(0)) wb.dispose();
});

async function fixture() {
  const store = createSpreadsheetStore();
  const wb = await WorkbookHandle.createDefault();
  live.push(wb);
  expect(wb.isStub).toBe(false);
  const history = new History();
  wb.attachHistory(history);
  wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
  wb.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
  wb.setNumber({ sheet: 0, row: 4, col: 1 }, 3);
  wb.setFormula({ sheet: 0, row: 0, col: 3 }, '=B2+1');
  wb.setFormula({ sheet: 0, row: 1, col: 3 }, '=B3*2');
  wb.setFormula({ sheet: 0, row: 2, col: 3 }, '=SUM(B2:C3)');
  wb.setFormula({ sheet: 0, row: 3, col: 3 }, '=B5+10');
  return { store, wb, history };
}

/** Cells, conditional formats and validations of sheet 0, ids included. */
function state(wb: WorkbookHandle) {
  wb.recalc();
  return {
    cells: Array.from(wb.physicalCells(0))
      .map((c) => ({ r: c.addr.row, c: c.addr.col, v: c.value, f: c.formula }))
      .sort((x, y) => x.r - y.r || x.c - y.c),
    cf: JSON.stringify(wb.getConditionalFormats(0)),
    dv: JSON.stringify(wb.engineWorkbook().sheet(0).listValidations()),
  };
}

const refusal = () => new FormulonError(StatusCode.InvalidArgument, 'refused', '');

/** Make the `nth` call across `methods` of the engine sheet 0 refuse. */
function refuseNth(wb: WorkbookHandle, methods: readonly string[], nth: number): () => number {
  const sheet = wb.engineWorkbook().sheet(0) as unknown as Record<
    string,
    (...a: unknown[]) => unknown
  >;
  let calls = 0;
  for (const method of methods) {
    const original = sheet[method]?.bind(sheet);
    if (!original) throw new Error(`engine sheet has no ${method}`);
    vi.spyOn(sheet, method).mockImplementation((...args: unknown[]) => {
      calls += 1;
      if (calls === nth) throw refusal();
      return original(...args);
    });
  }
  return () => calls;
}

/** State of the sheet right after the undo's inverse shift, before any restore step. */
function captureAfterInverseShift(wb: WorkbookHandle): () => ReturnType<typeof state> | null {
  const shift = wb.engineShiftCells.bind(wb);
  let after: ReturnType<typeof state> | null = null;
  vi.spyOn(wb, 'engineShiftCells').mockImplementation((range, direction) => {
    const ok = shift(range, direction);
    after = state(wb);
    return ok;
  });
  return () => after;
}

const expressionRule = (ranges: string) => ({
  type: CfRuleType.Expression,
  ranges,
  formula1: 'TRUE',
});

describe('cell shift restore under a refused write', () => {
  it('replay checkpoint restore leaves the cells as they were', async () => {
    const { store, wb } = await fixture();
    const checkpoint = captureShiftReplayCheckpoint(wb, store, BAND);
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 10);
    wb.setFormula({ sheet: 0, row: 0, col: 3 }, '=B2+5');
    wb.setFormula({ sheet: 0, row: 1, col: 3 }, '=B3*5');
    wb.setNumber({ sheet: 0, row: 6, col: 6 }, 7);
    const before = state(wb);

    const calls = refuseNth(wb, ['setFormula'], 2);
    expect(() => restoreShiftReplayCheckpoint(wb, store, BAND, checkpoint)).toThrow();
    expect(calls()).toBeGreaterThanOrEqual(2);
    expect(state(wb)).toEqual(before);
  });

  it('undo of a cell shift leaves the cells as the inverse shift left them', async () => {
    const { store, wb, history } = await fixture();
    expect(deleteCells(store, wb, history, BAND, 'up')).toBe(true);
    const afterShift = captureAfterInverseShift(wb);

    const calls = refuseNth(wb, ['setFormula'], 2);
    expect(() => history.undo()).toThrow();
    expect(calls()).toBeGreaterThanOrEqual(2);
    expect(afterShift()).not.toBeNull();
    expect(state(wb)).toEqual(afterShift());
    expect(history.canUndo()).toBe(true);
  });

  it('undo restores conditional formats and validations by id', async () => {
    const { store, wb, history } = await fixture();
    wb.engineWorkbook().sheet(0).addCfRule(expressionRule('B2:B6'));
    wb.engineWorkbook().sheet(0).addCfRule(expressionRule('C2:C6'));
    wb.engineWorkbook().sheet(0).addValidation({ type: 1, ranges: 'B2:B6', formula1: '0' });
    wb.engineWorkbook().sheet(0).addValidation({ type: 1, ranges: 'C2:C6', formula1: '0' });
    const before = state(wb);

    expect(deleteCells(store, wb, history, BAND, 'up')).toBe(true);
    expect(history.undo()).toBe(true);
    expect(state(wb).cf).toBe(before.cf);
    expect(state(wb).dv).toBe(before.dv);
  });

  it.each([
    { what: 'an updated conditional format', ranges: ['B2:B6', 'C2:C6'] },
    { what: 'a re-added conditional format', ranges: ['B2', 'C3'] },
  ])('undo refusing $what leaves the rules as the inverse shift left them', async ({ ranges }) => {
    const { store, wb, history } = await fixture();
    for (const range of ranges) wb.engineWorkbook().sheet(0).addCfRule(expressionRule(range));
    expect(deleteCells(store, wb, history, BAND, 'up')).toBe(true);
    const afterShift = captureAfterInverseShift(wb);

    const calls = refuseNth(wb, ['addCfRule', 'updateCfRule'], 2);
    expect(() => history.undo()).toThrow();
    expect(calls()).toBeGreaterThanOrEqual(2);
    expect(state(wb).cf).toBe(afterShift()?.cf);
  });

  it.each([
    { what: 'an updated validation', ranges: ['B2:B6', 'C2:C6'] },
    { what: 'a re-added validation', ranges: ['B2', 'C3'] },
  ])('undo refusing $what leaves the rules as the inverse shift left them', async ({ ranges }) => {
    const { store, wb, history } = await fixture();
    for (const range of ranges) {
      wb.engineWorkbook().sheet(0).addValidation({ type: 1, ranges: range, formula1: '0' });
    }
    expect(deleteCells(store, wb, history, BAND, 'up')).toBe(true);
    const afterShift = captureAfterInverseShift(wb);

    const calls = refuseNth(wb, ['addValidation', 'updateValidation'], 2);
    expect(() => history.undo()).toThrow();
    expect(calls()).toBeGreaterThanOrEqual(2);
    expect(state(wb).dv).toBe(afterShift()?.dv);
  });
});
