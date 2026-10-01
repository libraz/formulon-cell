// @vitest-environment node

import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { deleteRows, insertCols, insertRows } from '../../../src/commands/structure.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const merge = { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 };

describe('native row/column structure parity', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
  });

  it('moves merges and formulas through an insert/delete sequence', () => {
    mutators.mergeRange(store, merge);
    expect(wb.engineAddMerge(0, merge)).toBe(true);
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'anchor');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 2, col: 0 }, 2);
    wb.setNumber({ sheet: 0, row: 3, col: 0 }, 3);
    wb.setFormula({ sheet: 0, row: 6, col: 3 }, '=SUM(A2:A4)');
    wb.recalc();

    const history = new History();
    wb.attachHistory(history);
    insertRows(store, wb, history, 2, 1);
    expect(wb.getMerges(0)).toEqual([{ sheet: 0, r0: 1, c0: 0, r1: 4, c1: 1 }]);
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([
      { sheet: 0, r0: 1, c0: 0, r1: 4, c1: 1 },
    ]);
    expect(wb.cellFormula({ sheet: 0, row: 7, col: 3 })).toBe('=SUM(A2:A5)');

    deleteRows(store, wb, history, 1, 1);
    expect(wb.getMerges(0)).toEqual([{ sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 }]);
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([
      { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 },
    ]);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 3 })).toBe('=SUM(A2:A4)');

    history.undo();
    expect(wb.getMerges(0)).toEqual([{ sheet: 0, r0: 1, c0: 0, r1: 4, c1: 1 }]);
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([
      { sheet: 0, r0: 1, c0: 0, r1: 4, c1: 1 },
    ]);
    expect(wb.cellFormula({ sheet: 0, row: 7, col: 3 })).toBe('=SUM(A2:A5)');
    history.undo();
    expect(wb.getMerges(0)).toEqual([{ sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 }]);
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([merge]);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 1 });
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 3 })).toBe('=SUM(A2:A4)');

    history.redo();
    history.redo();
    expect(wb.getMerges(0)).toEqual([{ sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 }]);
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([merge]);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.cellFormula({ sheet: 0, row: 6, col: 3 })).toBe('=SUM(A2:A4)');
  });

  it('restores formulas on every sheet when a native delete is undone', () => {
    const secondSheet = wb.addSheet('Second');
    expect(secondSheet).toBe(1);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 7);
    wb.setError({ sheet: 0, row: 1, col: 1 }, 1);
    wb.setFormula({ sheet: 0, row: 0, col: 3 }, '=A2');
    wb.setFormula({ sheet: 1, row: 0, col: 3 }, '=Sheet1!A2');
    wb.recalc();

    const history = new History();
    wb.attachHistory(history);
    deleteRows(store, wb, history, 1, 1);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBe('=#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=#REF!');

    history.undo();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBe('=A2');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=Sheet1!A2');
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toMatchObject({ kind: 'error', code: 1 });

    history.redo();
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBe('=#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=#REF!');
  });

  it('drops a vertical merge when its final row is deleted', () => {
    const vertical = { sheet: 0, r0: 1, c0: 6, r1: 2, c1: 6 };
    mutators.mergeRange(store, vertical);
    expect(wb.engineAddMerge(0, vertical)).toBe(true);
    const history = new History();
    wb.attachHistory(history);
    deleteRows(store, wb, history, 2, 1);
    expect(wb.getMerges(0)).toEqual([]);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    history.undo();
    expect(wb.getMerges(0)).toEqual([vertical]);
    expect(Array.from(store.getState().merges.byAnchor.values())).toEqual([vertical]);
    history.redo();
    expect(wb.getMerges(0)).toEqual([]);
    expect(store.getState().merges.byAnchor.size).toBe(0);
  });

  it('inherits the preceding row and column format on native inserts', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontSize: 20, bold: true });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 0 },
      {
        fontSize: 9,
        hyperlink: 'https://example.test',
        comment: 'source',
        locked: false,
      },
    );
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        rowHeights: new Map([
          [0, 30],
          [1, 15],
        ]),
        colWidths: new Map([
          [3, 120],
          [4, 80],
        ]),
      },
    }));
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 3 }, { fontSize: 22 });

    insertRows(store, wb, null, 1, 1);
    const rowFormat = store.getState().format.formats.get('0:1:0');
    expect(rowFormat).toMatchObject({ fontSize: 20, bold: true });
    expect(rowFormat?.hyperlink).toBeUndefined();
    expect(rowFormat?.comment).toBeUndefined();
    expect(store.getState().layout.rowHeights.get(1)).toBe(30);

    insertCols(store, wb, null, 4, 1);
    expect(store.getState().format.formats.get('0:5:4')?.fontSize).toBe(22);
    expect(store.getState().layout.colWidths.get(4)).toBe(120);
  });
});
