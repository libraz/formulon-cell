import { describe, expect, it } from 'vitest';
import { createStubModule } from '../../../src/engine/stub-engine.js';
import type { Workbook } from '../../../src/engine/types.js';

const num = (wb: Workbook, sheet: number, row: number, col: number): number => {
  const r = wb.getValue(sheet, row, col);
  return r.value.kind === 1 ? r.value.number : Number.NaN;
};

const newWorkbook = (): Workbook => createStubModule().Workbook.createDefault();

describe('stub-engine recalc', () => {
  it('propagates a chain of formulas to a fixed point in one call', () => {
    const wb = newWorkbook();
    try {
      // A1 = 2; A2 = A1 * 3; A3 = A2 + 1
      wb.setNumber(0, 0, 0, 2);
      wb.setFormula(0, 1, 0, '=A1*3');
      wb.setFormula(0, 2, 0, '=A2+1');
      wb.recalc();
      expect(num(wb, 0, 1, 0)).toBe(6);
      expect(num(wb, 0, 2, 0)).toBe(7);
    } finally {
      wb.delete();
    }
  });

  it('converges regardless of cell insertion order', () => {
    // Insert the downstream formula first so its initial pass reads a stale
    // (zero) upstream. recalc must still settle to the correct value in one
    // call.
    const wb = newWorkbook();
    try {
      wb.setFormula(0, 2, 0, '=A2+1'); // downstream first
      wb.setFormula(0, 1, 0, '=A1*3');
      wb.setNumber(0, 0, 0, 2);
      wb.recalc();
      expect(num(wb, 0, 1, 0)).toBe(6);
      expect(num(wb, 0, 2, 0)).toBe(7);
    } finally {
      wb.delete();
    }
  });

  it('settles after upstream literal is replaced', () => {
    const wb = newWorkbook();
    try {
      wb.setNumber(0, 0, 0, 2);
      wb.setFormula(0, 1, 0, '=A1*3');
      wb.setFormula(0, 2, 0, '=A2+1');
      wb.recalc();
      expect(num(wb, 0, 2, 0)).toBe(7);
      // Replace upstream and recalc again.
      wb.setNumber(0, 0, 0, 5);
      wb.recalc();
      expect(num(wb, 0, 1, 0)).toBe(15);
      expect(num(wb, 0, 2, 0)).toBe(16);
    } finally {
      wb.delete();
    }
  });

  it('handles SUM over a range', () => {
    const wb = newWorkbook();
    try {
      wb.setNumber(0, 0, 0, 1);
      wb.setNumber(0, 1, 0, 2);
      wb.setNumber(0, 2, 0, 3);
      wb.setFormula(0, 3, 0, '=SUM(A1:A3)');
      wb.recalc();
      expect(num(wb, 0, 3, 0)).toBe(6);
    } finally {
      wb.delete();
    }
  });

  it('returns blank for unset cells', () => {
    const wb = newWorkbook();
    try {
      const r = wb.getValue(0, 5, 5);
      expect(r.status.ok).toBe(true);
      expect(r.value.kind).toBe(0);
    } finally {
      wb.delete();
    }
  });
});
