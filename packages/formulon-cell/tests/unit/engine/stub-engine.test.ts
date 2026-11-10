import { ValueKind } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { isFormulonError } from '../../../src/engine/result.js';
import { createStubModule } from '../../../src/engine/stub-engine.js';
import type { Workbook } from '../../../src/engine/types.js';

const num = (wb: Workbook, sheet: number, row: number, col: number): number => {
  const v = wb.sheet(sheet).getTypedValue(row, col);
  return v.kind === ValueKind.Number ? v.number : Number.NaN;
};

const newWorkbook = (): Workbook => createStubModule().Workbook.createDefault();

describe('stub-engine recalc', () => {
  it('propagates a chain of formulas to a fixed point in one call', () => {
    const wb = newWorkbook();
    try {
      // A1 = 2; A2 = A1 * 3; A3 = A2 + 1
      wb.sheet(0).setNumber(0, 0, 2);
      wb.sheet(0).setFormula(1, 0, '=A1*3');
      wb.sheet(0).setFormula(2, 0, '=A2+1');
      wb.recalc();
      expect(num(wb, 0, 1, 0)).toBe(6);
      expect(num(wb, 0, 2, 0)).toBe(7);
    } finally {
      wb.dispose();
    }
  });

  it('converges regardless of cell insertion order', () => {
    // Insert the downstream formula first so its initial pass reads a stale
    // (zero) upstream. recalc must still settle to the correct value in one
    // call.
    const wb = newWorkbook();
    try {
      wb.sheet(0).setFormula(2, 0, '=A2+1'); // downstream first
      wb.sheet(0).setFormula(1, 0, '=A1*3');
      wb.sheet(0).setNumber(0, 0, 2);
      wb.recalc();
      expect(num(wb, 0, 1, 0)).toBe(6);
      expect(num(wb, 0, 2, 0)).toBe(7);
    } finally {
      wb.dispose();
    }
  });

  it('settles after upstream literal is replaced', () => {
    const wb = newWorkbook();
    try {
      wb.sheet(0).setNumber(0, 0, 2);
      wb.sheet(0).setFormula(1, 0, '=A1*3');
      wb.sheet(0).setFormula(2, 0, '=A2+1');
      wb.recalc();
      expect(num(wb, 0, 2, 0)).toBe(7);
      // Replace upstream and recalc again.
      wb.sheet(0).setNumber(0, 0, 5);
      wb.recalc();
      expect(num(wb, 0, 1, 0)).toBe(15);
      expect(num(wb, 0, 2, 0)).toBe(16);
    } finally {
      wb.dispose();
    }
  });

  it('handles SUM over a range', () => {
    const wb = newWorkbook();
    try {
      wb.sheet(0).setNumber(0, 0, 1);
      wb.sheet(0).setNumber(1, 0, 2);
      wb.sheet(0).setNumber(2, 0, 3);
      wb.sheet(0).setFormula(3, 0, '=SUM(A1:A3)');
      wb.recalc();
      expect(num(wb, 0, 3, 0)).toBe(6);
    } finally {
      wb.dispose();
    }
  });

  it('returns blank for unset cells', () => {
    const wb = newWorkbook();
    try {
      const v = wb.sheet(0).getTypedValue(5, 5);
      expect(v.kind).toBe(ValueKind.Blank);
    } finally {
      wb.dispose();
    }
  });
});

describe('stub-engine defensive list handlers', () => {
  // Enumerations on the stub are plain lists that are empty rather than
  // throwing, so a caller iterating them never crashes the page.

  it('listDefinedNames returns an empty list instead of throwing', () => {
    const wb = newWorkbook();
    try {
      expect(wb.listDefinedNames()).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('listTables returns an empty list instead of throwing', () => {
    const wb = newWorkbook();
    try {
      expect(wb.sheet(0).listTables()).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('listPassthroughs returns an empty list instead of throwing', () => {
    const wb = newWorkbook();
    try {
      expect(wb.ooxml.listPassthroughs()).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('does NOT crash an iterate-the-list loop pattern', () => {
    const wb = newWorkbook();
    try {
      let crashed = false;
      try {
        for (const name of wb.listDefinedNames()) {
          if (!name.name) break;
        }
      } catch {
        crashed = true;
      }
      expect(crashed).toBe(false);
    } finally {
      wb.dispose();
    }
  });
});

describe('stub-engine sheet and cell counts', () => {
  it('reports the sheet count and refuses a missing sheet', () => {
    const wb = newWorkbook();
    try {
      expect(wb.sheetCount).toBe(1);
      expect(wb.sheet(0).listCells()).toHaveLength(0);
      expect(wb.hasSheet(9)).toBe(false);
      let error: unknown = null;
      try {
        wb.sheet(9);
      } catch (e) {
        error = e;
      }
      expect(isFormulonError(error)).toBe(true);
    } finally {
      wb.dispose();
    }
  });
});

describe('stub-engine physical-cell index', () => {
  class CountingMap extends Map<string, unknown> {
    keyEntries = 0;

    override keys(): MapIterator<string> {
      const iterator = super.keys();
      return {
        next: () => {
          const result = iterator.next();
          if (!result.done) this.keyEntries += 1;
          return result;
        },
        [Symbol.iterator]() {
          return this;
        },
        [Symbol.dispose]() {
          iterator.return?.();
        },
      } as MapIterator<string>;
    }
  }

  it('caches enumeration and invalidates only membership changes', () => {
    const wb = newWorkbook();
    try {
      for (let row = 0; row < 4; row += 1) wb.sheet(0).setNumber(row, 0, row + 1);
      const internals = wb as unknown as {
        sheets: Array<{ cells: Map<string, unknown>; cellIndex: unknown }>;
      };
      const sheet = internals.sheets[0];
      if (!sheet) throw new Error('missing stub sheet');
      const counted = new CountingMap(sheet.cells);
      sheet.cells = counted;

      for (let index = 0; index < 4; index += 1) {
        expect(wb.sheet(0).listCells()[index]).toBeDefined();
      }
      expect(counted.keyEntries).toBe(4);

      counted.keyEntries = 0;
      wb.sheet(0).setNumber(0, 0, 99);
      expect(wb.sheet(0).listCells()[0]?.value).toBe(99);
      expect(counted.keyEntries).toBe(0);

      wb.sheet(0).setNumber(4, 0, 5);
      counted.keyEntries = 0;
      expect(wb.sheet(0).listCells()[4]?.value).toBe(5);
      expect(counted.keyEntries).toBe(5);

      wb.sheet(0).setBlank(1, 0);
      counted.keyEntries = 0;
      const cells = wb.sheet(0).listCells();
      expect(cells[0]?.row).toBe(0);
      expect(cells[1]?.row).toBe(2);
      expect(counted.keyEntries).toBe(4);
    } finally {
      wb.dispose();
    }
  });
});
