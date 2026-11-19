import { CellShift, CfRuleType, PivotAggregation, PivotAxis, ValueKind } from '@libraz/formulon';
import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { workbookHandleOf } from '../../src/api/workbook.js';
import { deleteSheetAction } from '../../src/commands/sheet-actions.js';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];
const hosts: HTMLElement[] = [];

const mountSheet = async (): Promise<SpreadsheetInstance> => {
  const host = makeHost();
  hosts.push(host);
  const instance = await mount(host, { locale: 'en' });
  live.push(instance);
  return instance;
};

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
  hosts.length = 0;
});

const names = (instance: SpreadsheetInstance): string[] =>
  instance.workbook.listSheets().map((s) => s.name);

/** The labels of the sheet tabs the chrome shows. */
const tabLabels = (): string[] =>
  Array.from(
    (hosts[0] as HTMLElement).querySelectorAll<HTMLElement>(
      '.fc-host__sheetbar-tabs .fc-host__sheetbar-tab',
    ),
  ).map((tab) => tab.textContent?.trim() ?? '');

describe('facade journal: sheets', () => {
  it('addSheet undoes to the old tabs and redoes under the same name', async () => {
    const instance = await mountSheet();
    instance.workbook.addSheet('Extra');
    expect(names(instance)).toEqual(['Sheet1', 'Extra']);
    expect(tabLabels()).toEqual(['Sheet1', 'Extra']);

    instance.history.undo();
    expect(names(instance)).toEqual(['Sheet1']);
    expect(tabLabels()).toEqual(['Sheet1']);

    instance.history.redo();
    expect(names(instance)).toEqual(['Sheet1', 'Extra']);
    expect(tabLabels()).toEqual(['Sheet1', 'Extra']);
  });

  it('addSheet in front of another sheet keeps its position through undo and redo', async () => {
    const instance = await mountSheet();
    instance.workbook.addSheet('Last');
    instance.history.clear();
    instance.workbook.addSheet('First', { insertBefore: 0 });
    expect(names(instance)).toEqual(['First', 'Sheet1', 'Last']);
    instance.history.undo();
    expect(names(instance)).toEqual(['Sheet1', 'Last']);
    instance.history.redo();
    expect(names(instance)).toEqual(['First', 'Sheet1', 'Last']);
  });

  it('rename undoes and redoes, with the tab following', async () => {
    const instance = await mountSheet();
    instance.workbook.sheet(0).rename('Data');
    expect(tabLabels()).toEqual(['Data']);
    instance.history.undo();
    expect(names(instance)).toEqual(['Sheet1']);
    expect(tabLabels()).toEqual(['Sheet1']);
    instance.history.redo();
    expect(names(instance)).toEqual(['Data']);
    expect(tabLabels()).toEqual(['Data']);
  });

  it('two renames in one batch undo to the first name', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      s.rename('One');
      s.rename('Two');
    });
    instance.history.undo();
    expect(names(instance)).toEqual(['Sheet1']);
    instance.history.redo();
    expect(names(instance)).toEqual(['Two']);
  });

  it('move undoes and redoes, with the tabs following', async () => {
    const instance = await mountSheet();
    instance.workbook.addSheet('B');
    instance.workbook.addSheet('C');
    instance.history.clear();
    instance.workbook.sheet('Sheet1').move(null);
    expect(names(instance)).toEqual(['B', 'C', 'Sheet1']);
    expect(tabLabels()).toEqual(['B', 'C', 'Sheet1']);
    instance.history.undo();
    expect(names(instance)).toEqual(['Sheet1', 'B', 'C']);
    expect(tabLabels()).toEqual(['Sheet1', 'B', 'C']);
    instance.history.redo();
    expect(names(instance)).toEqual(['B', 'C', 'Sheet1']);
  });

  it('copy undoes to no copy and redoes under the same name', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setNumber('A1', 5);
    instance.history.clear();
    const copy = s.copy();
    expect(names(instance)).toEqual(['Sheet1', 'Sheet1 (2)']);
    expect(copy.getTypedValue('A1').kind).toBe(ValueKind.Number);
    instance.history.undo();
    expect(names(instance)).toEqual(['Sheet1']);
    expect(tabLabels()).toEqual(['Sheet1']);
    instance.history.redo();
    expect(names(instance)).toEqual(['Sheet1', 'Sheet1 (2)']);
    expect(instance.workbook.sheet('Sheet1 (2)').getTypedValue('A1').kind).toBe(ValueKind.Number);
  });

  it('a stale facade of an undone addSheet throws and the redo gets a new one', async () => {
    const instance = await mountSheet();
    const added = instance.workbook.addSheet('Extra');
    instance.history.undo();
    expect(() => added.name).toThrow();
    instance.history.redo();
    expect(instance.workbook.sheet('Extra').name).toBe('Extra');
  });

  it('a cell written before addSheet in one batch undoes with it in one step', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      s.setNumber('A1', 1);
      instance.workbook.addSheet('Extra');
      instance.workbook.sheet('Extra').setNumber('A1', 2);
    });
    instance.history.undo();
    expect(names(instance)).toEqual(['Sheet1']);
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);
    expect(instance.history.canUndo).toBe(false);
    instance.history.redo();
    expect(instance.workbook.sheet('Extra').getTypedValue('A1').kind).toBe(ValueKind.Number);
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Number);
  });

  it('removeSheet clears History', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setNumber('A1', 1);
    instance.workbook.addSheet('Extra');
    expect(instance.history.canUndo).toBe(true);
    instance.workbook.removeSheet('Extra');
    expect(instance.history.canUndo).toBe(false);
    expect(instance.history.canRedo).toBe(false);
  });

  it('removeSheet inside a batch with other writes clears History after the batch', async () => {
    const instance = await mountSheet();
    instance.workbook.addSheet('Extra');
    const s = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      s.setNumber('A1', 1);
      instance.workbook.removeSheet('Extra');
    });
    expect(instance.history.canUndo).toBe(false);
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Number);
  });

  it('UI Delete Sheet clears History so no entry replays onto another sheet', async () => {
    const instance = await mountSheet();
    const { store } = spreadsheetInternals(instance);
    instance.workbook.addSheet('Two');
    instance.workbook.sheet('Two').setNumber('A1', 9);
    expect(instance.history.canUndo).toBe(true);

    const handle = workbookHandleOf(instance.workbook);
    const ok = deleteSheetAction(
      { store, workbook: handle, history: handle.attachedHistory(), origin: 'instanceApi' },
      0,
    );
    expect(ok).toBe(true);
    expect(names(instance)).toEqual(['Two']);
    expect(instance.history.canUndo).toBe(false);
    instance.history.undo();
    expect(instance.workbook.sheet(0).getTypedValue('A1').kind).toBe(ValueKind.Number);
  });
});

describe('facade journal: structure', () => {
  it('insertRows undoes and redoes', async () => {
    const instance = await mountSheet();
    const { store } = spreadsheetInternals(instance);
    const s = instance.workbook.sheet(0);
    s.setNumber('A1', 1);
    s.setNumber('A2', 2);
    instance.history.clear();

    s.insertRows(0, 0);
    expect(s.getValue('A2')).toBe(1);
    expect(s.getValue('A3')).toBe(2);

    instance.history.undo();
    expect(s.getValue('A1')).toBe(1);
    expect(s.getValue('A2')).toBe(2);
    expect(s.getTypedValue('A3').kind).toBe(ValueKind.Blank);
    expect(store.getState().data.cells.size).toBe(2);

    instance.history.redo();
    expect(s.getValue('A2')).toBe(1);
    expect(s.getValue('A3')).toBe(2);
  });

  it('deleteRows undo restores the band cells, references, rules, merges and drawings', async () => {
    const instance = await mountSheet();
    const { store } = spreadsheetInternals(instance);
    const s = instance.workbook.sheet(0);
    const other = instance.workbook.addSheet('Other');
    s.setNumber('A1', 1);
    s.setNumber('A2', 20);
    s.setNumber('B2', 21);
    s.setNumber('A3', 30);
    s.setFormula('A4', '=A2+A3');
    other.setFormula('A1', '=Sheet1!A2');
    s.addCfRule({ type: CfRuleType.Expression, ranges: 'A2:A3', formula1: '=A2>1' });
    s.addMerge('A2:B2');
    const shape = s.addShape({ text: 'x', anchor: { fromRow: 1, fromCol: 3, toRow: 1, toCol: 4 } });
    instance.history.clear();

    s.deleteRows(1, 1);
    expect(s.getValue('A2')).toBe(30);
    expect(other.getFormula('A1')).toContain('#REF!');
    expect(s.listDrawings()).toHaveLength(0);
    expect(s.listMerges()).toHaveLength(0);

    instance.history.undo();
    expect(s.getValue('A1')).toBe(1);
    expect(s.getValue('A2')).toBe(20);
    expect(s.getValue('B2')).toBe(21);
    expect(s.getValue('A3')).toBe(30);
    expect(s.getFormula('A4')).toBe('=A2+A3');
    expect(other.getFormula('A1')).toBe('=Sheet1!A2');
    expect(s.listCfRules()).toHaveLength(1);
    expect(store.getState().conditional.rules).toHaveLength(1);
    expect(s.listMerges()).toHaveLength(1);
    expect(store.getState().merges.byAnchor.size).toBe(1);
    expect(s.listDrawings().map((d) => d.objectId)).toEqual([shape.objectId]);
    expect(store.getState().data.cells.get(`${0}:${1}:${0}`)).toBeDefined();

    instance.history.redo();
    expect(s.getValue('A2')).toBe(30);
    expect(s.listDrawings()).toHaveLength(0);
    expect(other.getFormula('A1')).toContain('#REF!');
  });

  it('deleteColumns undo restores the band cells', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setNumber('A1', 1);
    s.setNumber('B1', 2);
    s.setNumber('C1', 3);
    instance.history.clear();
    s.deleteColumns('B');
    expect(s.getValue('B1')).toBe(3);
    instance.history.undo();
    expect([s.getValue('A1'), s.getValue('B1'), s.getValue('C1')]).toEqual([1, 2, 3]);
    instance.history.redo();
    expect(s.getValue('B1')).toBe(3);
  });

  it('a batch of setNumber, insertRows and setNumber undoes in one step to the original grid', async () => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setNumber('A5', 1);
    instance.history.clear();

    instance.workbook.batch(() => {
      s.setNumber('A5', 2);
      s.insertRows(0, 0);
      s.setNumber('A5', 3);
    });
    expect(s.getValue('A5')).toBe(3);
    expect(s.getValue('A6')).toBe(2);

    instance.history.undo();
    expect(s.getValue('A5')).toBe(1);
    expect(s.getTypedValue('A6').kind).toBe(ValueKind.Blank);
    expect(instance.history.canUndo).toBe(false);

    instance.history.redo();
    expect(s.getValue('A5')).toBe(3);
    expect(s.getValue('A6')).toBe(2);
  });
});

describe('facade journal: cell shifts', () => {
  /** Values and XF indexes of `addrs`, as the engine holds them. */
  const snapshot = (instance: SpreadsheetInstance, addrs: string[]) => {
    const s = instance.workbook.sheet(0);
    return addrs.map((a) => [s.getValue(a), s.ooxml.getCellXfIndex(a)]);
  };

  const COLUMN = ['A1', 'A2', 'A3', 'A4'];
  const ROW = ['A1', 'B1', 'C1', 'D1'];

  /** A1 bold, then 2 and 3 along the column or the row. */
  const seed = async (along: string[]) => {
    const instance = await mountSheet();
    const s = instance.workbook.sheet(0);
    s.setNumber(along[0] as string, 1);
    s.setNumber(along[1] as string, 2);
    s.setNumber(along[2] as string, 3);
    s.updateCellFormat(along[0] as string, { font: { bold: true } });
    instance.history.clear();
    return { instance, s, before: snapshot(instance, along) };
  };

  const roundTrip = async (
    along: string[],
    write: (s: Awaited<ReturnType<typeof seed>>['s']) => void,
  ) => {
    const { instance, s, before } = await seed(along);
    write(s);
    const after = snapshot(instance, along);
    expect(after).not.toEqual(before);
    instance.history.undo();
    expect(snapshot(instance, along)).toEqual(before);
    instance.history.redo();
    expect(snapshot(instance, along)).toEqual(after);
  };

  it('insertCells shifting down undoes and redoes values and formats', async () => {
    await roundTrip(COLUMN, (s) => s.insertCells('A1', CellShift.Down));
  });

  it('insertCells shifting right undoes and redoes values and formats', async () => {
    await roundTrip(ROW, (s) => s.insertCells('A1', CellShift.Right));
  });

  it('deleteCells shifting up undoes and redoes the removed and the shifted cells', async () => {
    await roundTrip(COLUMN, (s) => s.deleteCells('A1', CellShift.Up));
  });

  it('deleteCells shifting left undoes and redoes the removed and the shifted cells', async () => {
    await roundTrip(ROW, (s) => s.deleteCells('A1', CellShift.Left));
  });
});

describe('facade journal: pivot details sheet', () => {
  it('showDetails undoes to no details sheet and redoes it', async () => {
    const instance = await mountSheet();
    const wb = instance.workbook;
    const s = wb.sheet(0);
    s.setValues('A1', [
      ['Region', 'Amount'],
      ['East', 10],
      ['West', 20],
    ]);
    const cache = wb.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:B3', sheet: s });
    cache.refresh();
    const pivot = s.addPivot({ name: 'P', cacheId: cache, anchor: 'D1' });
    pivot.setFieldAxis('Region', PivotAxis.Row);
    pivot.addDataField({ fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
    instance.history.clear();

    const details = pivot.showDetails('E2');
    expect(wb.listSheets()).toHaveLength(2);
    const added = details.name;
    expect(tabLabels()).toContain(added);

    instance.history.undo();
    expect(wb.listSheets().map((x) => x.name)).toEqual(['Sheet1']);
    expect(tabLabels()).toEqual(['Sheet1']);

    instance.history.redo();
    expect(wb.listSheets().map((x) => x.name)).toContain(added);
    expect(tabLabels()).toContain(added);
  });
});
