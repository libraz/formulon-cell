import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  createPivotTableFromRange,
  executeRibbonPivotTableAction,
} from '../../../src/commands/pivot-table.js';
import { PivotAxis } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** The measured fixture: Region / Product / Amt / Qty over five rows in A1:D6. */
const ROWS: (string | number)[][] = [
  ['Region', 'Product', 'Amt', 'Qty'],
  ['North', 'A', 100, 1],
  ['North', 'B', 200, 2],
  ['South', 'A', 300, 3],
  ['South', 'B', 400, 4],
  ['East', 'A', 500, 5],
];
const SOURCE = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 3 };

const strings = {
  pivotTable: 'PivotTable',
  pivotTableNewSheet: 'New sheet',
  pivotTableRefreshData: 'Refresh',
  pivotTableRefreshUnavailable: 'Unavailable',
  recommendedPivotTables: 'Recommended',
  pivotAuthoringDetail: 'Detail',
  workbookStructureProtectedBlocked: 'Protected',
};

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  open.push(wb);
  for (const [row, cells] of ROWS.entries()) {
    for (const [col, value] of cells.entries()) {
      if (typeof value === 'number') wb.setNumber({ sheet: 0, row, col }, value);
      else wb.setText({ sheet: 0, row, col }, value);
    }
  }
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
  return { wb, store, history: new History() };
};

const recommend = (deps: Awaited<ReturnType<typeof setup>>) =>
  executeRibbonPivotTableAction({
    origin: 'ribbon' as const,
    commandId: 'test',
    store: deps.store,
    workbook: deps.wb,
    action: 'recommended-insert',
    strings,
    history: deps.history,
  });

const axisNames = (wb: WorkbookHandle, sheet: number, axis: number): string[] => {
  const definition = wb.getPivotDefinition(sheet, 0);
  const order = axis === PivotAxis.Row ? definition?.rowFieldOrder : definition?.colFieldOrder;
  return (order ?? []).map((index) => definition?.fields[index]?.sourceName ?? '');
};

// Mac desktop, as measured on this one fixture: no dialog; a new sheet before
// the source sheet with the first text field on rows and the first number
// field on columns, and no values.
describe('Recommended PivotTables on the Mac profile', () => {
  it('inserts the report on a new sheet before the source, at A3, as one undo step', async () => {
    const deps = await setup();
    const { wb, store, history } = deps;
    const sheets = wb.sheetCount;
    const result = recommend(deps);
    expect(result).toMatchObject({ kind: 'created', destination: { sheet: 0, row: 2, col: 0 } });
    expect(wb.sheetCount).toBe(sheets + 1);
    expect(axisNames(wb, 0, PivotAxis.Row)).toEqual(['Region']);
    expect(axisNames(wb, 0, PivotAxis.Col)).toEqual(['Amt']);
    expect(wb.getPivotDefinition(0, 0)?.dataFields).toEqual([]);
    expect(store.getState().data.sheetIndex).toBe(0);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 0 });
    // The report's cells reach the sheet the user now sees.
    const shown = [...store.getState().data.cells.values()].map((cell) => cell.value);
    expect(shown).toContainEqual({ kind: 'text', value: 'East' });

    expect(history.undo()).toBe(true);
    expect(wb.sheetCount).toBe(sheets);
  });

  it('builds the report on an existing cache of the same source', async () => {
    const deps = await setup();
    const { wb } = deps;
    const existing = createPivotTableFromRange(wb, {
      source: SOURCE,
      destination: { sheet: 0, row: 10, col: 6 },
      rowField: 'Region',
      valueField: 'Amt',
    });
    if (!existing.ok) throw new Error('fixture pivot failed');
    expect(recommend(deps)).toMatchObject({ kind: 'created' });
    expect(wb.getPivotDefinition(0, 0)?.cacheId).toBe(existing.cacheId);
  });

  it('shares the cache when the source sheet name reads as a cell reference', async () => {
    const deps = await setup();
    const { wb } = deps;
    expect(wb.renameSheet(0, 'Q1')).toBe(true);
    const existing = createPivotTableFromRange(wb, {
      source: SOURCE,
      destination: { sheet: 0, row: 10, col: 6 },
      rowField: 'Region',
      valueField: 'Amt',
    });
    if (!existing.ok) throw new Error('fixture pivot failed');
    expect(recommend(deps)).toMatchObject({ kind: 'created' });
    expect(wb.getPivotDefinition(0, 0)?.cacheId).toBe(existing.cacheId);
  });

  it('builds a new cache when none has the source', async () => {
    const deps = await setup();
    const { wb } = deps;
    expect(recommend(deps)).toMatchObject({ kind: 'created' });
    expect(wb.getPivotDefinition(0, 0)?.cacheId).toBeGreaterThanOrEqual(0);
  });

  it('keeps the plain Recommended action as it was', async () => {
    const deps = await setup();
    const result = executeRibbonPivotTableAction({
      origin: 'ribbon' as const,
      commandId: 'test',
      store: deps.store,
      workbook: deps.wb,
      action: 'recommended',
      strings,
      history: deps.history,
    });
    expect(result.kind).toBe('report');
  });
});
