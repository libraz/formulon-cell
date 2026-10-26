import { AnchorEditAs } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { recordSlicersChange } from '../../../src/commands/slice-history.js';
import {
  createSlicer,
  recomputeSlicerFilters,
  removeSlicer,
  setSlicerSelected,
  updateSlicer,
} from '../../../src/commands/slicers.js';
import { deleteRows, insertRows } from '../../../src/commands/structure.js';
import { hydrateAutoFilterFromEngine } from '../../../src/engine/auto-filter-sync.js';
import { drawingAnchorOf } from '../../../src/engine/drawing-geometry.js';
import { hydrateSlicersFromEngine } from '../../../src/engine/slicer-sync.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const GRID: string[][] = [
  ['Item', 'Region', 'Qty'],
  ['A', 'East', '1'],
  ['B', 'West', '2'],
  ['C', 'East', '3'],
  ['D', 'West', '4'],
];
const table: Range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 };

async function bind(wb: WorkbookHandle) {
  open.push(wb);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  hydrateTableOverlaysFromEngine(wb, store);
  hydrateAutoFilterFromEngine(wb, store, 0);
  hydrateSlicersFromEngine(wb, store);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { wb, store, history };
}

/** A workbook holding Table1 over A1:C5 with its store bound. */
async function session(grid: string[][] = GRID) {
  const wb = await WorkbookHandle.createDefault();
  grid.forEach((cols, row) => {
    cols.forEach((text, col) => {
      wb.setText({ sheet: 0, row, col }, text);
    });
  });
  const bound = await bind(wb);
  formatAsTable(bound.store, table, { workbook: wb });
  return bound;
}

const reopen = async (wb: WorkbookHandle) => bind(await WorkbookHandle.loadBytes(wb.save()));

const only = (store: SpreadsheetStore) => {
  const slicers = store.getState().slicers.slicers;
  expect(slicers).toHaveLength(1);
  return slicers[0] as (typeof slicers)[number];
};

const add = (store: SpreadsheetStore, wb: WorkbookHandle, extra = {}) => {
  const result = createSlicer(store, wb, {
    tableName: 'Table1',
    column: 'Region',
    x: 320,
    y: 120,
    ...extra,
  });
  if (!result.ok) throw new Error(result.reason);
  return result.spec;
};

describe('table slicers saved with the workbook', () => {
  it('writes a slicer and hydrates it from the saved file', async () => {
    const { wb, store } = await session();
    add(store, wb, { caption: 'Pick a region', style: 'SlicerStyleLight2', columnCount: 2 });

    expect(wb.listSlicers(0)).toHaveLength(1);
    expect(wb.listSlicers(0)[0]).toMatchObject({
      tableName: 'Table1',
      columnName: 'Region',
      caption: 'Pick a region',
      style: 'SlicerStyleLight2',
      columnCount: 2,
    });
    expect(only(store).native?.name).toBe(wb.listSlicers(0)[0]?.name);

    const reopened = await reopen(wb);
    expect(only(reopened.store)).toMatchObject({
      tableName: 'Table1',
      column: 'Region',
      caption: 'Pick a region',
      style: 'SlicerStyleLight2',
      columnCount: 2,
      sheet: 0,
    });
  });

  it('reads the selection back from the table filter', async () => {
    const { wb, store } = await session();
    const spec = add(store, wb, { selected: ['East'] });
    expect(recomputeSlicerFilters(store, wb)).toBe(2);
    expect(wb.getTableAutoFilter(0)?.columns[0]).toMatchObject({ colId: 1, values: ['East'] });

    const reopened = await reopen(wb);
    expect(only(reopened.store).selected).toEqual(['East']);

    setSlicerSelected(store, spec.id, []);
    recomputeSlicerFilters(store, wb);
    expect(only((await reopen(wb)).store).selected).toEqual([]);
  });

  it('moves the slicer by its anchor and updates caption, style and columns', async () => {
    const { wb, store } = await session();
    const spec = add(store, wb);
    const before = wb.listSlicers(0)[0];

    updateSlicer(store, spec.id, { x: 700, y: 400, caption: 'Regions', style: '', columnCount: 3 });

    const after = wb.listSlicers(0)[0];
    expect(after).toMatchObject({ caption: 'Regions', style: '', columnCount: 3 });
    expect([after?.fromRow, after?.fromCol]).not.toEqual([before?.fromRow, before?.fromCol]);
    // The panel keeps its size while it moves.
    expect((after?.toRow ?? 0) - (after?.fromRow ?? 0)).toBe(
      (before?.toRow ?? 0) - (before?.fromRow ?? 0),
    );
    const reopened = await reopen(wb);
    expect(only(reopened.store)).toMatchObject({ caption: 'Regions', columnCount: 3 });
    expect(
      drawingAnchorOf({ ...after, editAs: AnchorEditAs.TwoCell, cx: 0, cy: 0 } as never),
    ).toMatchObject({
      fromRow: after?.fromRow,
    });
  });

  it('removes the slicer from the file, and undo adds it back with its name', async () => {
    const { wb, store, history } = await session();
    const spec = add(store, wb, { caption: 'Regions' });
    const name = wb.listSlicers(0)[0]?.name;

    recordSlicersChange(history, store, () => removeSlicer(store, spec.id));
    expect(wb.listSlicers(0)).toEqual([]);
    expect((await reopen(wb)).store.getState().slicers.slicers).toEqual([]);

    history.undo();
    expect(wb.listSlicers(0)).toHaveLength(1);
    expect(wb.listSlicers(0)[0]).toMatchObject({ name, caption: 'Regions' });
    history.redo();
    expect(wb.listSlicers(0)).toEqual([]);
  });

  it('undoes and redoes a new slicer in the file', async () => {
    const { wb, store, history } = await session();
    recordSlicersChange(history, store, () => {
      add(store, wb);
    });
    expect(wb.listSlicers(0)).toHaveLength(1);

    history.undo();
    expect(wb.listSlicers(0)).toEqual([]);
    history.redo();
    expect(wb.listSlicers(0)).toHaveLength(1);
  });

  it('shifts the slicer once when rows are inserted, and back on undo', async () => {
    const { wb, store, history } = await session();
    add(store, wb);
    const row = wb.listSlicers(0)[0]?.fromRow ?? 0;

    expect(insertRows(store, wb, history, 0, 2)).toBe(true);
    expect(wb.listSlicers(0)[0]?.fromRow).toBe(row + 2);
    expect(wb.listSlicers(0)).toHaveLength(1);

    history.undo();
    expect(wb.listSlicers(0)[0]?.fromRow).toBe(row);
    history.redo();
    expect(wb.listSlicers(0)[0]?.fromRow).toBe(row + 2);
  });

  it('puts a slicer back where it was when a row delete over it is undone', async () => {
    const { wb, store, history } = await session();
    add(store, wb, { style: 'SlicerStyleDark1' });
    const before = wb.listSlicers(0)[0] as { fromRow: number; toRow: number; name: string };

    expect(
      deleteRows(store, wb, history, before.fromRow - 1, before.toRow - before.fromRow + 2),
    ).toBe(true);
    expect(wb.listSlicers(0)).toHaveLength(1);
    expect(wb.listSlicers(0)[0]?.toRow).toBe(wb.listSlicers(0)[0]?.fromRow);

    history.undo();
    expect(wb.listSlicers(0)).toHaveLength(1);
    expect(wb.listSlicers(0)[0]).toMatchObject({
      name: before.name,
      style: 'SlicerStyleDark1',
      fromRow: before.fromRow,
      toRow: before.toRow,
    });
    expect(store.getState().slicers.slicers).toHaveLength(1);
  });

  it('gives slicers with non-ASCII names distinct, stable ids', async () => {
    const { wb, store } = await session([
      ['品目', '地域', '数量'],
      ['A', '東', '1'],
      ['B', '西', '2'],
    ]);
    for (const column of ['地域', '品目']) add(store, wb, { column });

    const reopened = await reopen(wb);
    const ids = reopened.store.getState().slicers.slicers.map((spec) => spec.id);
    expect(new Set(ids).size).toBe(2);
    expect(ids).toEqual((await reopen(wb)).store.getState().slicers.slicers.map((spec) => spec.id));
  });

  it('ties a slicer to its anchor and reads the loaded size', async () => {
    const { wb, store } = await session();
    add(store, wb);
    const native = wb.listSlicers(0)[0];

    expect(only(store).nativeAnchor).toMatchObject({
      fromRow: native?.fromRow,
      toRow: native?.toRow,
    });
    const loaded = only((await reopen(wb)).store);
    expect(loaded.w).toBeGreaterThan(0);
    expect(loaded.nativeAnchor?.toCol).toBe(native?.toCol);
  });
});
