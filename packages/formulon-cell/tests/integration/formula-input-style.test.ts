import { describe, expect, it } from 'vitest';
import { pasteTSV } from '../../src/commands/clipboard/paste.js';
import { pasteSpecial } from '../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../src/commands/clipboard/snapshot.js';
import { fillRange } from '../../src/commands/fill.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../src/store/store.js';

const D4 = { sheet: 0, row: 3, col: 3 } as const;
const E4 = { sheet: 0, row: 3, col: 4 } as const;
const F4 = { sheet: 0, row: 3, col: 5 } as const;
const source = { sheet: 0, r0: D4.row, c0: D4.col, r1: D4.row, c1: D4.col } as const;

describe('integration: formula input reference style', () => {
  it.each([false, true])(
    'coerces an external TSV formula using the active UI style (%s)',
    async (r1c1) => {
      const workbook = await WorkbookHandle.createDefault();
      try {
        expect(workbook.isStub).toBe(false);
        const store = createSpreadsheetStore();
        mutators.setR1C1(store, r1c1);
        mutators.setActive(store, D4);
        workbook.setNumber(E4, 19);

        expect(pasteTSV(store.getState(), workbook, '=RC5')).not.toBeNull();
        expect(workbook.cellFormula(D4)).toBe(r1c1 ? '=$E4' : '=RC5');
        if (r1c1) expect(workbook.getValue(D4)).toEqual({ kind: 'number', value: 19 });
      } finally {
        workbook.dispose();
      }
    },
  );

  it('does not reinterpret canonical formulas during internal fill and snapshot paste', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      const store = createSpreadsheetStore();
      mutators.setR1C1(store, true);
      workbook.setFormula(D4, '=RC5');
      mutators.replaceCells(store, workbook.cells(0));

      expect(fillRange(store.getState(), workbook, source, { ...source, c1: E4.col })).toBe(true);
      expect(workbook.cellFormula(E4)).toBe('=RD5');

      const snapshot = captureSnapshot(store.getState(), source);
      expect(snapshot).not.toBeNull();
      mutators.setActive(store, F4);
      if (snapshot) {
        pasteSpecial(store.getState(), store, workbook, snapshot, {
          what: 'all',
          operation: 'none',
          skipBlanks: false,
          transpose: false,
        });
      }
      expect(workbook.cellFormula(F4)).toBe('=RE5');
    } finally {
      workbook.dispose();
    }
  });
});
