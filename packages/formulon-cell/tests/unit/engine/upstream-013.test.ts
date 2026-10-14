import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const addr = (row: number) => ({ sheet: 0, row, col: 0 });

describe('formulon 0.13 integration', () => {
  it.each([1, 101, 9, 109])(
    'excludes nested SUBTOTAL cells for code %i through save/load',
    async (code) => {
      const first = await WorkbookHandle.createDefault();
      try {
        expect(first.isStub).toBe(false);
        expect(first.version).toBe('0.15.2');
        first.withBatchedRecalc(() => {
          for (let row = 0; row < 3; row += 1) first.setNumber(addr(row), 0);
          first.setFormula(addr(3), `=SUBTOTAL(${code},A1:A3)`);
          first.setNumber(addr(4), 100);
          first.setFormula(addr(5), `=SUBTOTAL(${code},A5:A5)`);
          first.setFormula(addr(6), `=SUBTOTAL(${code},A1:A6)`);
        });
        const expected = code === 1 || code === 101 ? 25 : 100;
        expect(first.getValue(addr(6))).toEqual({ kind: 'number', value: expected });
        const loaded = await WorkbookHandle.loadBytes(first.save());
        try {
          expect(loaded.getValue(addr(6))).toEqual({ kind: 'number', value: expected });
          loaded.setNumber(addr(4), 200);
          expect(loaded.getValue(addr(6))).toEqual({ kind: 'number', value: expected * 2 });
        } finally {
          loaded.dispose();
        }
      } finally {
        first.dispose();
      }
    },
  );
});
