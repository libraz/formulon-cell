import { describe, expect, it } from 'vitest';
import { findMatchingCells } from '../../../src/commands/goto-special.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('native range and formula reads', () => {
  it('pages sparse cells without omissions and connects selection search', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.rangeEnumeration).toBe(true);
      wb.withBatchedRecalc(() => {
        for (const row of [0, 3, 8, 40, 100]) wb.setNumber({ sheet: 0, row, col: 2 }, row + 1);
        wb.setNumber({ sheet: 0, row: 8, col: 8 }, 999);
      });
      const range = { sheet: 0, r0: 3, c0: 2, r1: 100, c1: 2 };
      const entries = Array.from(wb.cellsInRange(range, 2));
      expect(entries.map((entry) => entry.addr.row)).toEqual([3, 8, 40, 100]);
      expect(entries.map((entry) => entry.value)).toEqual(
        [4, 9, 41, 101].map((value) => ({ kind: 'number', value })),
      );
      const store = createSpreadsheetStore();
      mutators.setRange(store, range);
      expect(
        findMatchingCells(wb, store, 'selection', 'constants').map((addr) => addr.row),
      ).toEqual([3, 8, 40, 100]);
      expect(() => Array.from(wb.cellsInRange(range, 0))).toThrow(RangeError);
    } finally {
      wb.dispose();
    }
  });

  it('keeps blank physical cells eligible for validation and conditional-format searches', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const addr = { sheet: 0, row: 2, col: 2 };
      wb.setNumber(addr, 1);
      wb.setBlank(addr);
      const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 4 };
      expect(Array.from(wb.cellsInRange(range))).toEqual([]);
      const store = createSpreadsheetStore();
      mutators.setRange(store, range);
      mutators.setCellFormat(store, addr, { validation: { kind: 'list', source: ['a'] } });
      mutators.addConditionalRule(store, { kind: 'data-bar', range, color: '#638ec6' });
      expect(findMatchingCells(wb, store, 'selection', 'data-validation')).toEqual([addr]);
      expect(findMatchingCells(wb, store, 'selection', 'conditional-format')).toEqual([addr]);
    } finally {
      wb.dispose();
    }
  });

  it('reads external workbook references in native R1C1 form', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const addr = { sheet: 0, row: 0, col: 0 };
      wb.setFormula(addr, '=[Book.xlsx]Sheet!A1');
      expect(wb.cellFormula(addr)).toBe('=[Book.xlsx]Sheet!A1');
      expect(wb.cellFormulaR1C1(addr)).toBe('=[Book.xlsx]Sheet!RC');
    } finally {
      wb.dispose();
    }
  });

  it('reads A1 and relative R1C1 formulas before and after save', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const addr = { sheet: 0, row: 2, col: 2 };
      wb.setFormula(addr, '=A1+$B$2');
      expect(wb.cellFormula(addr)).toBe('=A1+$B$2');
      expect(wb.cellFormulaR1C1(addr)).toBe('=R[-2]C[-2]+R2C2');
      expect(wb.cellFormulaR1C1({ ...addr, row: 1 })).toBeNull();
      const loaded = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(loaded.cellFormulaR1C1(addr)).toBe('=R[-2]C[-2]+R2C2');
      } finally {
        loaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
