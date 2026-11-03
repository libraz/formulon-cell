import { describe, expect, it } from 'vitest';
import { pasteSpecial } from '../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { setActive } from './paste-special/fixtures.js';

describe('pasting a dynamic-array spill as Excel measures it', () => {
  it('re-creates the spill from its anchor: I1:I3 =SEQUENCE(3) pasted at K1', async () => {
    const wb = await WorkbookHandle.createDefault();
    if (!wb.capabilities.spillInfo) return;
    const store = createSpreadsheetStore();
    wb.setFormula({ sheet: 0, row: 0, col: 8 }, '=SEQUENCE(3)');
    wb.recalc();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      for (const row of [0, 1, 2]) {
        const addr = { sheet: 0, row, col: 8 };
        cells.set(addrKey(addr), {
          value: wb.getValue(addr),
          formula: row === 0 ? '=SEQUENCE(3)' : null,
        });
      }
      return { ...s, data: { ...s.data, cells } };
    });
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 8, r1: 2, c1: 8 });
    if (!snap) throw new Error('no snapshot');
    setActive(store, 0, 10);
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 10 })).toBe('=SEQUENCE(3)');
    expect(wb.spillInfo(0, 2, 10)).toMatchObject({ anchorRow: 0, anchorCol: 10, rows: 3 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 10 })).toEqual({ kind: 'number', value: 3 });
  });
});
