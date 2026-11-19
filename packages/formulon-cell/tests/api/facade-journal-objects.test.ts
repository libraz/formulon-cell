import { ChartPlotType, PivotAggregation, PivotAxis, ValueKind } from '@libraz/formulon';
import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];

const mountSheet = async (): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en' });
  live.push(instance);
  return instance;
};

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

/** A source block of a header row and two data rows at A1:B3. */
const seedSource = (instance: SpreadsheetInstance): void => {
  instance.workbook.sheet(0).setValues('A1', [
    ['Region', 'Amount'],
    ['East', 10],
    ['West', 20],
  ]);
};

describe('facade journal: objects', () => {
  it('addTable undoes to no table and blank header cells, and redoes', async () => {
    const instance = await mountSheet();
    const { store } = spreadsheetInternals(instance);
    const s = instance.workbook.sheet(0);
    s.addTable({ range: 'A1:B3', name: 'Scores' });
    expect(s.listTables()).toHaveLength(1);
    expect(store.getState().tables.tables).toHaveLength(1);
    expect(s.getTypedValue('A1').kind).not.toBe(ValueKind.Blank);

    instance.history.undo();
    expect(s.listTables()).toHaveLength(0);
    expect(store.getState().tables.tables).toHaveLength(0);
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);

    instance.history.redo();
    expect(s.listTables().map((t) => t.name)).toEqual(['Scores']);
    expect(store.getState().tables.tables).toHaveLength(1);
    expect(s.getTypedValue('A1').kind).not.toBe(ValueKind.Blank);
  });

  it('a cell outside every table written in the same batch as addTable round-trips', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      s.setNumber('F10', 7);
      s.addTable({ range: 'A1:B3', name: 'Scores' });
    });
    expect(s.getTypedValue('F10').kind).toBe(ValueKind.Number);
    instance.history.undo();
    expect(s.listTables()).toHaveLength(0);
    expect(s.getTypedValue('F10').kind).toBe(ValueKind.Blank);
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);
    instance.history.redo();
    expect(s.listTables()).toHaveLength(1);
    expect(s.getTypedValue('F10').kind).toBe(ValueKind.Number);
  });

  it('a table rename undoes and redoes', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    const table = s.addTable({ range: 'A1:B3', name: 'Scores' });
    instance.history.clear();
    table.rename('Totals');
    expect(s.listTables().map((t) => t.name)).toEqual(['Totals']);
    instance.history.undo();
    expect(s.listTables().map((t) => t.name)).toEqual(['Scores']);
    instance.history.redo();
    expect(s.listTables().map((t) => t.name)).toEqual(['Totals']);
  });

  it('table styles undo and redo', async () => {
    const instance = await mountSheet();
    const wb = instance.workbook;
    const before = wb.listTableStyles().length;
    wb.setTableStyle({ name: 'Mine' } as Parameters<typeof wb.setTableStyle>[0]);
    expect(wb.listTableStyles()).toHaveLength(before + 1);
    instance.history.undo();
    expect(wb.listTableStyles()).toHaveLength(before);
    instance.history.redo();
    expect(wb.listTableStyles()).toHaveLength(before + 1);
  });

  it('removeDrawing undoes to the same id and z-order', async () => {
    const instance = await mountSheet();
    const { store } = spreadsheetInternals(instance);
    const s = instance.workbook.sheet(0);
    const anchor = { fromRow: 5, fromCol: 5, toRow: 8, toCol: 8 };
    const ids = [1, 2, 3].map((i) => s.addShape({ text: `s${i}`, anchor }).objectId);
    const order = (): number[] => s.listDrawings().map((d) => d.objectId);
    expect(order()).toEqual(ids);
    instance.history.clear();

    s.removeDrawing(ids[1] as number);
    expect(order()).toEqual([ids[0], ids[2]]);
    instance.history.undo();
    expect(order()).toEqual(ids);
    expect(store.getState().drawings.order).toHaveLength(3);
    instance.history.redo();
    expect(order()).toEqual([ids[0], ids[2]]);
    expect(store.getState().drawings.order).toHaveLength(2);
  });

  it('Chart.setTitle undoes and redoes', async () => {
    const instance = await mountSheet();
    const { store } = spreadsheetInternals(instance);
    seedSource(instance);
    const s = instance.workbook.sheet(0);
    const chart = s.addChart({ type: ChartPlotType.Bar, range: 'A1:B3' });
    instance.history.clear();
    const title = (): unknown => s.listCharts()[0]?.getProperties().title;
    const before = JSON.stringify(title());
    chart.setTitle('Revenue');
    const after = JSON.stringify(title());
    expect(after).not.toBe(before);
    const stored = (): string => JSON.stringify(store.getState().charts.charts);

    const storedAfter = stored();
    instance.history.undo();
    expect(JSON.stringify(title())).toBe(before);
    expect(stored()).not.toBe(storedAfter);
    instance.history.redo();
    expect(JSON.stringify(title())).toBe(after);
    expect(stored()).toBe(storedAfter);
  });

  describe('pivots', () => {
    const setup = (instance: SpreadsheetInstance) => {
      seedSource(instance);
      const wb = instance.workbook;
      const cache = wb.addPivotCache();
      cache.setWorksheetSource({ range: 'A1:B3', sheet: wb.sheet(0) });
      cache.refresh();
      instance.history.clear();
      return cache;
    };

    it('addPivot undoes to no pivot and redoes', async () => {
      const instance = await mountSheet();
      const s = instance.workbook.sheet(0);
      const cache = setup(instance);
      s.addPivot({ name: 'P', cacheId: cache, anchor: 'D1' });
      expect(s.listPivots()).toHaveLength(1);
      instance.history.undo();
      expect(s.listPivots()).toHaveLength(0);
      instance.history.redo();
      expect(s.listPivots().map((p) => p.name)).toEqual(['P']);
    });

    it('removePivot undoes to the pivot with its fields and redoes', async () => {
      const instance = await mountSheet();
      const s = instance.workbook.sheet(0);
      const cache = setup(instance);
      const pivot = s.addPivot({ name: 'P', cacheId: cache, anchor: 'D1' });
      pivot.setFieldAxis('Region', PivotAxis.Row);
      pivot.addDataField({ fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
      instance.history.clear();
      const fields = (): number => s.listPivots()[0]?.listDataFields().length ?? -1;
      expect(fields()).toBe(1);

      s.removePivot(pivot);
      expect(s.listPivots()).toHaveLength(0);
      instance.history.undo();
      expect(s.listPivots().map((p) => p.name)).toEqual(['P']);
      expect(fields()).toBe(1);
      instance.history.redo();
      expect(s.listPivots()).toHaveLength(0);
    });

    it('addPivotCache and addPivot in a batch are one step', async () => {
      const instance = await mountSheet();
      const wb = instance.workbook;
      const s = wb.sheet(0);
      seedSource(instance);
      instance.history.clear();
      wb.batch(() => {
        const cache = wb.addPivotCache();
        s.addPivot({ name: 'P', cacheId: cache, anchor: 'D1' });
      });
      expect(wb.listPivotCaches()).toHaveLength(1);
      expect(s.listPivots()).toHaveLength(1);
      instance.history.undo();
      expect(wb.listPivotCaches()).toHaveLength(0);
      expect(s.listPivots()).toHaveLength(0);
      expect(instance.history.canUndo).toBe(false);
      instance.history.redo();
      expect(wb.listPivotCaches()).toHaveLength(1);
      expect(s.listPivots()).toHaveLength(1);
    });

    it('addPivotCache undoes to no cache and redoes', async () => {
      const instance = await mountSheet();
      const wb = instance.workbook;
      wb.addPivotCache();
      instance.history.undo();
      expect(wb.listPivotCaches()).toHaveLength(0);
      instance.history.redo();
      expect(wb.listPivotCaches()).toHaveLength(1);
    });

    it('cache.refresh clears History', async () => {
      const instance = await mountSheet();
      const cache = setup(instance);
      instance.workbook.sheet(0).setNumber('C1', 1);
      expect(instance.history.canUndo).toBe(true);
      cache.refresh();
      expect(instance.history.canUndo).toBe(false);
    });

    it('setWorksheetSource and removePivotCache clear History', async () => {
      const instance = await mountSheet();
      const wb = instance.workbook;
      const cache = setup(instance);
      wb.sheet(0).setNumber('C1', 1);
      cache.setWorksheetSource({ range: 'A1:B3', sheet: wb.sheet(0) });
      expect(instance.history.canUndo).toBe(false);
      wb.sheet(0).setNumber('C1', 2);
      wb.removePivotCache(cache);
      expect(instance.history.canUndo).toBe(false);
    });
  });

  it('addDataTable undoes to no record and redoes', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setValues('A1', [
      [null, '=E1*2'],
      [1, null],
      [2, null],
    ]);
    s.setNumber('E1', 5);
    instance.history.clear();
    s.addDataTable({ range: 'B2:B3', columnInput: 'E1' });
    expect(s.getDataTable('B2')).not.toBeNull();
    instance.history.undo();
    expect(s.getDataTable('B2')).toBeNull();
    instance.history.redo();
    expect(s.getDataTable('B2')).not.toBeNull();
  });

  it('removeDataTable undoes to the record and redoes', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setValues('A1', [
      [null, '=E1*2'],
      [1, null],
      [2, null],
    ]);
    s.setNumber('E1', 5);
    s.addDataTable({ range: 'B2:B3', columnInput: 'E1' });
    instance.history.clear();
    s.removeDataTable('B2');
    expect(s.getDataTable('B2')).toBeNull();
    instance.history.undo();
    expect(s.getDataTable('B2')).not.toBeNull();
    instance.history.redo();
    expect(s.getDataTable('B2')).toBeNull();
  });

  it('reports the result cells of addDataTable in the change event', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setValues('A1', [
      [null, '=E1*2'],
      [1, null],
      [2, null],
    ]);
    s.setNumber('E1', 5);
    await Promise.resolve();
    const seen: string[] = [];
    const parts: string[][] = [];
    instance.workbook.on('change', (e) => {
      for (const c of e.cells) seen.push(`${c.at.row}:${c.at.col}`);
      parts.push([...e.parts]);
    });
    s.addDataTable({ range: 'B2:B3', columnInput: 'E1' });
    await Promise.resolve();
    expect(seen).toEqual(expect.arrayContaining(['1:1', '2:1']));
    // A data table is a cell-side record, like an array formula: it names no other part.
    expect(parts).toEqual([[]]);
  });

  it('refreshes the store result cells when a removed data table is undone', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    const { store } = spreadsheetInternals(instance);
    s.setValues('A1', [
      [null, '=E1*2'],
      [1, null],
      [2, null],
    ]);
    s.setNumber('E1', 5);
    s.addDataTable({ range: 'B2:B3', columnInput: 'E1' });
    const resultOf = () => store.getState().data.cells.get('0:1:1')?.value;
    const withTable = resultOf();
    expect(withTable).toEqual({ kind: 'number', value: 2 });
    s.removeDataTable('B2');
    expect(resultOf()).not.toEqual(withTable);
    instance.history.undo();
    expect(resultOf()).toEqual(withTable);
  });
});
