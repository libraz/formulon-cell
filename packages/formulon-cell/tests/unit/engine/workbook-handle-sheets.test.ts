import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ChangeEvent } from '../../../src/engine/workbook-handle.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

/**
 * Unit: workbook-handle sheet ops against the stub engine. The class is the
 * single entry point for engine state; this suite locks in its sheet-mutate
 * surface (add / rename / remove / move) and the change-event contract that
 * the chrome layer subscribes to.
 */
describe('WorkbookHandle — sheet mutations', () => {
  let wb: WorkbookHandle;
  let events: ChangeEvent[];
  let unsub: () => void;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    events = [];
    unsub = wb.subscribe((e) => {
      events.push(e);
    });
  });

  afterEach(() => {
    unsub();
    try {
      wb.dispose();
    } catch {
      // dispose may already have run inside the test; we just need teardown
      // to never throw across the suite boundary.
    }
  });

  it('starts with at least one sheet', () => {
    expect(wb.sheetCount).toBeGreaterThanOrEqual(1);
    expect(wb.sheetName(0)).toMatch(/^Sheet\d+$/);
  });

  it('addSheet appends and emits sheet-add', () => {
    const before = wb.sheetCount;
    const idx = wb.addSheet('Data');
    expect(idx).toBe(before);
    expect(wb.sheetCount).toBe(before + 1);
    expect(wb.sheetName(idx)).toBe('Data');
    expect(events.at(-1)).toEqual({ kind: 'sheet-add', index: idx, name: 'Data' });
  });

  it('addSheet without a name uses an auto-generated unique name', () => {
    const idx1 = wb.addSheet();
    const idx2 = wb.addSheet();
    expect(wb.sheetName(idx1)).not.toBe(wb.sheetName(idx2));
    expect(wb.sheetName(idx1)).toMatch(/^Sheet\d+$/);
  });

  it('renameSheet updates the name and emits sheet-rename (when capability supports it)', () => {
    if (!wb.capabilities.sheetMutate) {
      // Stub engine may not expose rename; the API contract is that an
      // unsupported capability returns false without throwing.
      expect(wb.renameSheet(0, 'Renamed')).toBe(false);
      return;
    }
    const ok = wb.renameSheet(0, 'Renamed');
    expect(ok).toBe(true);
    expect(wb.sheetName(0)).toBe('Renamed');
    expect(events.at(-1)).toEqual({ kind: 'sheet-rename', index: 0, name: 'Renamed' });
  });

  it('removeSheet drops the sheet and emits sheet-remove (when capability supports it)', () => {
    wb.addSheet('Extra');
    const before = wb.sheetCount;
    if (!wb.capabilities.sheetMutate) {
      expect(wb.removeSheet(before - 1)).toBe(false);
      return;
    }
    const ok = wb.removeSheet(before - 1);
    expect(ok).toBe(true);
    expect(wb.sheetCount).toBe(before - 1);
    expect(events.at(-1)).toEqual({ kind: 'sheet-remove', index: before - 1 });
  });

  it('subscribe disposer stops further events', () => {
    const fn = vi.fn();
    const localUnsub = wb.subscribe(fn);
    wb.addSheet('A');
    expect(fn).toHaveBeenCalledTimes(1);
    localUnsub();
    wb.addSheet('B');
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('post-dispose calls throw via assertAlive', () => {
    wb.dispose();
    expect(() => wb.sheetCount).toThrow();
  });
});

describe('WorkbookHandle — values and formulas', () => {
  let wb: WorkbookHandle;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault({ preferStub: true });
  });

  afterEach(() => wb.dispose());

  it('setNumber / getValue round-trips', () => {
    const a1 = { sheet: 0, row: 0, col: 0 };
    wb.setNumber(a1, 42);
    expect(wb.getValue(a1)).toEqual({ kind: 'number', value: 42 });
  });

  it('setText / getValue round-trips', () => {
    const b2 = { sheet: 0, row: 1, col: 1 };
    wb.setText(b2, 'hello');
    expect(wb.getValue(b2)).toEqual({ kind: 'text', value: 'hello' });
  });

  it('setBlank clears a populated cell', () => {
    const a1 = { sheet: 0, row: 0, col: 0 };
    wb.setNumber(a1, 7);
    wb.setBlank(a1);
    expect(wb.getValue(a1)).toEqual({ kind: 'blank' });
  });

  it('setFormula records the formula text and evaluates', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 3);
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 4);
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=A1+B1');
    const v = wb.getValue({ sheet: 0, row: 0, col: 2 });
    expect(v.kind === 'number' ? v.value : null).toBe(7);
  });
});
