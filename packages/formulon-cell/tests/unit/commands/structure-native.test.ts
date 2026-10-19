// @vitest-environment node

import { AnchorEditAs, AnchorKind, ChartInsertKind } from '@libraz/formulon';
import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../../../src/commands/structure.js';
import { hydrateChartsFromEngine } from '../../../src/engine/charts-sync.js';
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
  it('restores the charts a row delete removes and re-reads the ones it moves', () => {
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'v');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'a');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'b');
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
    const insert = (row: number, heightEmu: number): number => {
      const id = wb.insertChart(0, {
        kind: ChartInsertKind.Column,
        source: 'Sheet1!$A$1:$B$3',
        placement: {
          anchorKind: AnchorKind.TwoCell,
          editAs: AnchorEditAs.TwoCell,
          row,
          col: 4,
          widthEmu: 1_905_000,
          heightEmu,
        },
      });
      if (id === null) throw new Error('chart was not inserted');
      return id;
    };
    // Wholly inside rows 10..14, so the delete removes it; the other starts below.
    const inside = insert(10, 380_000);
    const below = insert(20, 380_000);
    hydrateChartsFromEngine(wb, store, 0);
    const charts = (): { id: number; fromRow: number }[] =>
      store.getState().charts.charts.map((chart) => ({
        id: chart.nativeObjectId,
        fromRow: chart.nativeAnchor.fromRow,
      }));
    expect(charts()).toEqual([
      { id: inside, fromRow: 10 },
      { id: below, fromRow: 20 },
    ]);

    const history = new History();
    wb.attachHistory(history);
    deleteRows(store, wb, history, 10, 5);
    expect(charts()).toEqual([{ id: below, fromRow: 15 }]);

    history.undo();
    expect(charts()).toEqual([
      { id: inside, fromRow: 10 },
      { id: below, fromRow: 20 },
    ]);
    expect(wb.getChart(0, inside)?.objectId).toBe(inside);

    history.redo();
    expect(charts()).toEqual([{ id: below, fromRow: 15 }]);
  });

  it('keeps a chart a column delete only cuts through', () => {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    const id = wb.insertChart(0, {
      kind: ChartInsertKind.Line,
      source: 'Sheet1!$A$1:$A$2',
      placement: { anchorKind: AnchorKind.TwoCell, row: 3, col: 2 },
    });
    if (id === null) throw new Error('chart was not inserted');
    hydrateChartsFromEngine(wb, store, 0);
    const history = new History();
    wb.attachHistory(history);
    deleteCols(store, wb, history, 3, 1);
    expect(store.getState().charts.charts.map((chart) => chart.nativeObjectId)).toEqual([id]);
    history.undo();
    expect(store.getState().charts.charts.map((chart) => chart.nativeObjectId)).toEqual([id]);
  });
});
