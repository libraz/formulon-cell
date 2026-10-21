import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  createPivotTableAtPlacement,
  executeRibbonPivotTableAction,
  pivotSourceForSelection,
} from '../../../src/commands/pivot-table.js';
import { setWorkbookStructureProtected } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Sales'],
    ['East', 12],
    ['West', 5],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r + 1, col: c + 1 };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  return { wb, store };
};

const names = (wb: WorkbookHandle): string[] =>
  Array.from({ length: wb.sheetCount }, (_, i) => wb.sheetName(i));

const strings = {
  pivotTable: 'PivotTable',
  pivotTableNewSheet: 'New sheet',
  pivotTableRefreshData: 'Refresh',
  pivotTableRefreshUnavailable: 'Unavailable',
  recommendedPivotTables: 'Recommended',
  pivotAuthoringDetail: 'Detail',
  workbookStructureProtectedBlocked: 'Protected',
  dataFieldNaming: ja.pivotTableDialog.dataFieldNaming,
};

describe('pivot source for a single cell', () => {
  it('expands to the current region', async () => {
    const { wb, store } = await setup();
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });
    expect(pivotSourceForSelection(store, wb)).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 });
    wb.dispose();
  });

  it('keeps a multi-cell selection and an empty cell as selected', async () => {
    const { wb, store } = await setup();
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    expect(pivotSourceForSelection(store, wb)).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    mutators.setActive(store, { sheet: 0, row: 20, col: 20 });
    expect(pivotSourceForSelection(store, wb)).toEqual({
      sheet: 0,
      r0: 20,
      c0: 20,
      r1: 20,
      c1: 20,
    });
    wb.dispose();
  });
});

describe('new-sheet placement', () => {
  it('inserts the sheet before the source, anchors at A3 and undoes as one step', async () => {
    const { wb, store } = await setup();
    const history = new History();
    const result = createPivotTableAtPlacement(
      { store, workbook: wb, history },
      {
        source: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 },
        rowField: 'Region',
        valueField: 'Sales',
      },
      'new-sheet',
    );
    expect(result).toMatchObject({ ok: true, destination: { sheet: 0, row: 2, col: 0 } });
    expect(names(wb)).toEqual(['Sheet2', 'Sheet1']);
    expect(wb.getPivotTables()).toMatchObject([{ sheetIndex: 0, top: 2, left: 0 }]);
    expect(wb.getValue({ sheet: 1, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'Region' });

    expect(history.undo()).toBe(true);
    expect(names(wb)).toEqual(['Sheet1']);
    expect(wb.getPivotTables()).toEqual([]);

    expect(history.redo()).toBe(true);
    expect(names(wb)).toEqual(['Sheet2', 'Sheet1']);
    expect(wb.getPivotTables()).toMatchObject([{ sheetIndex: 0, top: 2, left: 0 }]);
    wb.dispose();
  });

  it('undoes a report placed on the source sheet', async () => {
    const { wb, store } = await setup();
    const history = new History();
    const result = createPivotTableAtPlacement(
      { store, workbook: wb, history },
      {
        source: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 },
        rowField: 'Region',
        valueField: 'Sales',
      },
      { row: 8, col: 1 },
    );
    expect(result.ok).toBe(true);
    expect(wb.getPivotTables()).toHaveLength(1);
    history.undo();
    expect(wb.getPivotTables()).toEqual([]);
    wb.dispose();
  });

  it('removes the new sheet again when the report cannot be built', async () => {
    const { wb, store } = await setup();
    const history = new History();
    const result = createPivotTableAtPlacement(
      { store, workbook: wb, history },
      {
        source: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 },
        rowField: 'Missing',
        valueField: 'Sales',
      },
      'new-sheet',
    );
    expect(result).toMatchObject({ ok: false, reason: 'invalid-field' });
    expect(names(wb)).toEqual(['Sheet1']);
    expect(history.canUndo()).toBe(false);
    wb.dispose();
  });
});

describe('ribbon one-click', () => {
  it('lands on the same default destination as the dialog', async () => {
    const { wb, store } = await setup();
    mutators.setActive(store, { sheet: 0, row: 2, col: 1 });
    const result = executeRibbonPivotTableAction({
      store,
      workbook: wb,
      action: 'new-sheet',
      strings,
      history: new History(),
    });
    expect(result).toMatchObject({
      kind: 'created',
      destinationSheet: 0,
      destination: { sheet: 0, row: 2, col: 0 },
    });
    expect(names(wb)).toEqual(['Sheet2', 'Sheet1']);
    expect([...wb.pivotCells(0)].map((c) => c.value)).toContainEqual({
      kind: 'text',
      value: '合計 / Sales',
    });
    wb.dispose();
  });

  it('reports a protected workbook structure', async () => {
    const { wb, store } = await setup();
    setWorkbookStructureProtected(store, true);
    mutators.setActive(store, { sheet: 0, row: 2, col: 1 });
    const result = executeRibbonPivotTableAction({
      store,
      workbook: wb,
      action: 'new-sheet',
      strings,
    });
    expect(result).toMatchObject({ kind: 'report' });
    expect(names(wb)).toEqual(['Sheet1']);
    wb.dispose();
  });
});
