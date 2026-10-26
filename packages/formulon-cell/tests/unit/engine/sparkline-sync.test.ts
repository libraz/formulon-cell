import { SparklineType } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearSparkline,
  listSparklines,
  setSparkline,
  setSparklines,
  sparklineAt,
} from '../../../src/commands/sparkline.js';
import { deleteRows, insertCols, insertRows } from '../../../src/commands/structure.js';
import { hydrateSparklinesFromEngine } from '../../../src/engine/sparkline-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  type Sparkline,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const at = (row: number, col: number) => ({ sheet: 0, row, col });

/** A workbook with data in B2:C5, its store hydrated and bound. */
async function session(seed?: (wb: WorkbookHandle) => void, bytes?: Uint8Array) {
  const wb = bytes ? await WorkbookHandle.loadBytes(bytes) : await WorkbookHandle.createDefault();
  open.push(wb);
  if (!bytes) {
    for (let row = 1; row <= 4; row += 1) {
      wb.setNumber(at(row, 1), row);
      wb.setNumber(at(row, 2), row * 2 - 5);
    }
  }
  seed?.(wb);
  const store = createSpreadsheetStore();
  hydrateSparklinesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { wb, store, history };
}

const saved = async (wb: WorkbookHandle) => session(undefined, wb.save());
const groups = (wb: WorkbookHandle) => wb.listSparklineGroups(0);
const line: Sparkline = { kind: 'line', source: 'B2:B5', color: '#cc3300' };

describe('sparklines saved with the workbook', () => {
  it('writes a sparkline as a group and hydrates it from the saved file', async () => {
    const { wb, store } = await session();
    setSparkline(store, at(1, 4), { ...line, showNegative: true, negativeColor: '#112233' });

    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]).toMatchObject({
      type: SparklineType.Line,
      negative: true,
      sparklines: [{ formula: 'Sheet1!B2:B5', location: 'E2' }],
    });
    expect((groups(wb)[0]?.colorSeries.rgb ?? 0) & 0xffffff).toBe(0xcc3300);

    const reopened = await saved(wb);
    expect(sparklineAt(reopened.store.getState(), at(1, 4))).toMatchObject({
      kind: 'line',
      source: 'B2:B5',
      color: '#cc3300',
      showNegative: true,
      negativeColor: '#112233',
    });
  });

  it('maps column and win-loss to engine group types', async () => {
    const { wb, store } = await session();
    setSparkline(store, at(1, 4), { kind: 'column', source: 'B2:B5', showNegative: true });
    setSparkline(store, at(2, 4), { kind: 'win-loss', source: 'C2:C5', showNegative: true });

    expect(groups(wb).map((group) => group.type)).toEqual([
      SparklineType.Column,
      SparklineType.Stacked,
    ]);
    const reopened = await saved(wb);
    expect(listSparklines(reopened.store.getState()).map((entry) => entry.spec.kind)).toEqual([
      'column',
      'win-loss',
    ]);
  });

  it('keeps sparklines created together in one group', async () => {
    const { wb, store } = await session();
    setSparklines(store, [
      { addr: at(1, 4), spec: line },
      { addr: at(2, 4), spec: { ...line, source: 'C2:C5' } },
    ]);

    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]?.sparklines.map((s) => s.location)).toEqual(['E2', 'E3']);
    const reopened = await saved(wb);
    const [first, second] = listSparklines(reopened.store.getState());
    expect(first?.spec.group).toBeDefined();
    expect(first?.spec.group).toBe(second?.spec.group);

    clearSparkline(store, at(1, 4));
    expect(groups(wb)[0]?.sparklines.map((s) => s.location)).toEqual(['E3']);
    clearSparkline(store, at(2, 4));
    expect(groups(wb)).toEqual([]);
  });

  it('edits the colour and kind of a group and nothing else', async () => {
    const { wb, store } = await session((handle) => {
      handle.addSparklineGroup(0, {
        type: SparklineType.Line,
        markers: true,
        high: true,
        lineWeight: 2.25,
        displayXAxis: true,
        sparklines: [{ formula: 'Sheet1!B2:B5', location: 'E2' }],
      });
    });
    const spec = sparklineAt(store.getState(), at(1, 4)) as Sparkline;
    expect(spec.nativeGroup).toMatchObject({ markers: true, high: true, lineWeight: 2.25 });

    setSparkline(store, at(1, 4), { ...spec, color: '#00aa00' }, undefined);

    const [group] = groups(wb);
    expect((group?.colorSeries.rgb ?? 0) & 0xffffff).toBe(0x00aa00);
    expect(group).toMatchObject({
      markers: true,
      high: true,
      lineWeight: 2.25,
      displayXAxis: true,
    });
    expect(groups(wb)).toHaveLength(1);
  });

  it('puts a cleared group back with its settings on undo', async () => {
    const { wb, store, history } = await session((handle) => {
      handle.addSparklineGroup(0, {
        type: SparklineType.Column,
        markers: true,
        lineWeight: 3,
        sparklines: [{ formula: 'Sheet1!B2:B5', location: 'E2' }],
      });
    });
    clearSparkline(store, at(1, 4), history);
    expect(groups(wb)).toEqual([]);

    history.undo();
    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]).toMatchObject({
      type: SparklineType.Column,
      markers: true,
      lineWeight: 3,
    });
    history.redo();
    expect(groups(wb)).toEqual([]);
  });

  it('undoes and redoes a new sparkline in the file', async () => {
    const { wb, store, history } = await session();
    setSparkline(store, at(1, 4), line, history);
    expect(groups(wb)).toHaveLength(1);

    history.undo();
    expect(groups(wb)).toEqual([]);
    history.redo();
    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]?.sparklines[0]?.location).toBe('E2');
  });

  const rowOf = (store: SpreadsheetStore): number =>
    listSparklines(store.getState())[0]?.addr.row ?? -1;

  it('shifts a sparkline once when rows or columns are inserted, and back on undo', async () => {
    const { wb, store, history } = await session();
    setSparkline(store, at(1, 4), line);

    expect(insertRows(store, wb, history, 0, 2)).toBe(true);
    expect(groups(wb)[0]?.sparklines).toEqual([{ formula: 'Sheet1!B4:B7', location: 'E4' }]);
    expect(rowOf(store)).toBe(3);
    expect(sparklineAt(store.getState(), at(3, 4))?.source).toBe('B4:B7');

    expect(insertCols(store, wb, history, 0, 1)).toBe(true);
    expect(groups(wb)[0]?.sparklines).toEqual([{ formula: 'Sheet1!C4:C7', location: 'F4' }]);

    history.undo();
    history.undo();
    expect(groups(wb)[0]?.sparklines).toEqual([{ formula: 'Sheet1!B2:B5', location: 'E2' }]);
    expect(rowOf(store)).toBe(1);
    history.redo();
    expect(groups(wb)[0]?.sparklines).toEqual([{ formula: 'Sheet1!B4:B7', location: 'E4' }]);
  });

  it('restores a sparkline whose cell a row delete removed', async () => {
    const { wb, store, history } = await session();
    setSparkline(store, at(1, 4), line, history);

    expect(deleteRows(store, wb, history, 1, 1)).toBe(true);
    expect(groups(wb)).toEqual([]);
    expect(listSparklines(store.getState())).toEqual([]);

    history.undo();
    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]?.sparklines[0]).toMatchObject({ location: 'E2' });
    expect(listSparklines(store.getState())).toHaveLength(1);
  });

  it('keeps a sparkline on its sheet when an earlier sheet is removed', async () => {
    const { wb, store } = await session();
    const second = wb.addSheet();
    setSparkline(store, { sheet: second, row: 1, col: 4 }, { kind: 'line', source: 'B2:B5' });
    expect(wb.listSparklineGroups(second)).toHaveLength(1);

    expect(wb.removeSheet(0)).toBe(true);

    expect(wb.listSparklineGroups(0)).toHaveLength(1);
    expect(listSparklines(store.getState()).map((entry) => entry.addr.sheet)).toEqual([0]);
  });

  it('applies a style change of one member to the whole group', async () => {
    const { wb, store } = await session();
    setSparklines(store, [
      { addr: at(1, 4), spec: line },
      { addr: at(2, 4), spec: { ...line, source: 'C2:C5' } },
    ]);
    const second = sparklineAt(store.getState(), at(2, 4)) as Sparkline;

    setSparkline(store, at(2, 4), { ...second, color: '#00aa00', group: second.group });

    expect((groups(wb)[0]?.colorSeries.rgb ?? 0) & 0xffffff).toBe(0x00aa00);
    expect(sparklineAt(store.getState(), at(1, 4))?.color).toBe('#00aa00');
    expect(sparklineAt(store.getState(), at(2, 4))?.color).toBe('#00aa00');
  });

  it('re-reads the group from the engine when it refuses a member change', async () => {
    const { wb, store } = await session();
    const members = vi.spyOn(wb, 'setSparklineGroupSparklines').mockReturnValue(false);
    setSparklines(store, [
      { addr: at(1, 4), spec: line },
      { addr: at(2, 4), spec: line },
    ]);
    members.mockRestore();

    expect(listSparklines(store.getState()).map((entry) => entry.addr.row)).toEqual([1]);
    expect(groups(wb)[0]?.sparklines).toHaveLength(1);
  });
});
