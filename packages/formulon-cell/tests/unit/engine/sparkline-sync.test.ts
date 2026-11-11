import { BlanksAs, SparklineAxisType, SparklineType } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearSparkline,
  defaultSparklineGroupSettings,
  groupOfCell,
  groupSparklines,
  listSparklines,
  rgbColorSpec,
  setSparkline,
  setSparklines,
  sparklineAt,
  ungroupSparklines,
  updateSparklineGroup,
} from '../../../src/commands/sparkline.js';
import { deleteRows, insertCols, insertRows } from '../../../src/commands/structure.js';
import { formatA1Cell, formatA1Range } from '../../../src/engine/address.js';
import { hydrateSparklinesFromEngine } from '../../../src/engine/sparkline-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SparklineGroupSettings,
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
/** Sheet 0's groups, each member as its source and location text. */
const groups = (wb: WorkbookHandle) =>
  wb.listSparklineGroups(0).map((group) => ({
    ...group,
    sparklines: group.sparklines.map((member) => ({
      formula: member.source
        ? `${wb.sheetName(member.source.sheet)}!${formatA1Range({
            r0: member.source.rect.firstRow,
            c0: member.source.rect.firstCol,
            r1: member.source.rect.lastRow,
            c1: member.source.rect.lastCol,
          })}`
        : null,
      location: formatA1Cell(member.location.row, member.location.col),
    })),
  }));
const lineSettings = {
  ...defaultSparklineGroupSettings(SparklineType.Line),
  colorSeries: rgbColorSpec('#cc3300'),
};
const SOURCE = 'B2:B5';
const rgbOf = (color: { argb: number } | null | undefined): number => (color?.argb ?? 0) & 0xffffff;

describe('sparklines saved with the workbook', () => {
  it('writes a sparkline as a group and hydrates it from the saved file', async () => {
    const { wb, store } = await session();
    setSparkline(store, at(1, 4), SOURCE, {
      ...lineSettings,
      negative: true,
      colorNegative: rgbColorSpec('#112233'),
    });

    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]).toMatchObject({
      type: SparklineType.Line,
      negative: true,
      sparklines: [{ formula: 'Sheet1!B2:B5', location: 'E2' }],
    });
    expect(rgbOf(groups(wb)[0]?.colorSeries)).toBe(0xcc3300);

    const reopened = await saved(wb);
    const state = reopened.store.getState();
    expect(sparklineAt(state, at(1, 4))).toMatchObject({ source: SOURCE });
    const settings = groupOfCell(state, at(1, 4))?.settings;
    expect(settings).toMatchObject({ type: SparklineType.Line, negative: true });
    expect(rgbOf(settings?.colorSeries)).toBe(0xcc3300);
    expect(rgbOf(settings?.colorNegative)).toBe(0x112233);
  });

  it('maps column and win-loss to engine group types', async () => {
    const { wb, store } = await session();
    setSparkline(store, at(1, 4), SOURCE, defaultSparklineGroupSettings(SparklineType.Column));
    setSparkline(store, at(2, 4), 'C2:C5', defaultSparklineGroupSettings(SparklineType.Stacked));

    expect(groups(wb).map((group) => group.type)).toEqual([
      SparklineType.Column,
      SparklineType.Stacked,
    ]);
    const reopened = await saved(wb);
    expect(
      listSparklines(reopened.store.getState()).map(
        (entry) => groupOfCell(reopened.store.getState(), entry.addr)?.settings.type,
      ),
    ).toEqual([SparklineType.Column, SparklineType.Stacked]);
  });

  it('keeps sparklines created together in one group', async () => {
    const { wb, store } = await session();
    setSparklines(
      store,
      [
        { addr: at(1, 4), source: SOURCE },
        { addr: at(2, 4), source: 'C2:C5' },
      ],
      lineSettings,
    );

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
    expect(store.getState().sparkline.groups.size).toBe(0);
  });

  it('round-trips every group setting through the engine', async () => {
    const { wb, store } = await session();
    const settings: SparklineGroupSettings = {
      type: SparklineType.Column,
      colorSeries: rgbColorSpec('#102030'),
      colorNegative: rgbColorSpec('#112233'),
      colorAxis: rgbColorSpec('#223344'),
      colorMarkers: rgbColorSpec('#334455'),
      colorFirst: rgbColorSpec('#445566'),
      colorLast: rgbColorSpec('#556677'),
      colorHigh: rgbColorSpec('#667788'),
      colorLow: rgbColorSpec('#778899'),
      markers: true,
      high: true,
      low: true,
      first: true,
      last: true,
      negative: true,
      displayXAxis: true,
      displayHidden: true,
      minAxisType: SparklineAxisType.Custom,
      maxAxisType: SparklineAxisType.Group,
      manualMin: -4.5,
      manualMax: 12,
      dateAxis: false,
      dateAxisRange: null,
      lineWidthPt: 2.25,
      emptyCellsAs: BlanksAs.Span,
    };
    setSparkline(store, at(1, 4), SOURCE, settings);

    const reopened = await saved(wb);
    const read = groupOfCell(reopened.store.getState(), at(1, 4))?.settings;
    expect(read).toBeDefined();
    const colors = (g: SparklineGroupSettings | undefined) =>
      [
        g?.colorSeries,
        g?.colorNegative,
        g?.colorAxis,
        g?.colorMarkers,
        g?.colorFirst,
        g?.colorLast,
        g?.colorHigh,
        g?.colorLow,
      ].map(rgbOf);
    expect(colors(read)).toEqual(colors(settings));
    expect(read).toMatchObject({
      type: SparklineType.Column,
      markers: true,
      high: true,
      low: true,
      first: true,
      last: true,
      negative: true,
      displayXAxis: true,
      displayHidden: true,
      minAxisType: SparklineAxisType.Custom,
      maxAxisType: SparklineAxisType.Group,
      manualMin: -4.5,
      manualMax: 12,
      lineWidthPt: 2.25,
      emptyCellsAs: BlanksAs.Span,
    });
  });

  it('edits the colour of a group and nothing else', async () => {
    const { wb, store } = await session((handle) => {
      handle.addSparklineGroup(0, {
        type: SparklineType.Line,
        markers: true,
        high: true,
        lineWidthPt: 2.25,
        displayXAxis: true,
        sparklines: [{ source: 'Sheet1!B2:B5', location: 'E2' }],
      });
    });
    const group = groupOfCell(store.getState(), at(1, 4));
    expect(group?.settings).toMatchObject({ markers: true, high: true, lineWidthPt: 2.25 });

    updateSparklineGroup(store, group?.id as string, { colorSeries: rgbColorSpec('#00aa00') });

    const [engine] = groups(wb);
    expect(rgbOf(engine?.colorSeries)).toBe(0x00aa00);
    expect(engine).toMatchObject({
      markers: true,
      high: true,
      lineWidthPt: 2.25,
      displayXAxis: true,
    });
    expect(groups(wb)).toHaveLength(1);
  });

  it('puts a cleared group back with its settings on undo', async () => {
    const { wb, store, history } = await session((handle) => {
      handle.addSparklineGroup(0, {
        type: SparklineType.Column,
        markers: true,
        lineWidthPt: 3,
        sparklines: [{ source: 'Sheet1!B2:B5', location: 'E2' }],
      });
    });
    clearSparkline(store, at(1, 4), history);
    expect(groups(wb)).toEqual([]);

    history.undo();
    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]).toMatchObject({
      type: SparklineType.Column,
      markers: true,
      lineWidthPt: 3,
    });
    history.redo();
    expect(groups(wb)).toEqual([]);
  });

  it('undoes and redoes a new sparkline in the file', async () => {
    const { wb, store, history } = await session();
    setSparkline(store, at(1, 4), SOURCE, lineSettings, history);
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
    setSparkline(store, at(1, 4), SOURCE, lineSettings);

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
    setSparkline(store, at(1, 4), SOURCE, lineSettings, history);

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
    setSparkline(store, { sheet: second, row: 1, col: 4 }, SOURCE, lineSettings);
    expect(wb.listSparklineGroups(second)).toHaveLength(1);

    expect(wb.removeSheet(0)).toBe(true);

    expect(wb.listSparklineGroups(0)).toHaveLength(1);
    expect(listSparklines(store.getState()).map((entry) => entry.addr.sheet)).toEqual([0]);
  });

  it('applies a settings change to the whole group', async () => {
    const { wb, store } = await session();
    setSparklines(
      store,
      [
        { addr: at(1, 4), source: SOURCE },
        { addr: at(2, 4), source: 'C2:C5' },
      ],
      lineSettings,
    );
    const group = groupOfCell(store.getState(), at(2, 4));

    updateSparklineGroup(store, group?.id as string, { colorSeries: rgbColorSpec('#00aa00') });

    expect(rgbOf(groups(wb)[0]?.colorSeries)).toBe(0x00aa00);
    expect(groups(wb)).toHaveLength(1);
    expect(groupOfCell(store.getState(), at(1, 4))?.id).toBe(group?.id);
  });

  it('writes grouping and ungrouping as engine groups', async () => {
    const { wb, store } = await session();
    setSparkline(store, at(1, 4), SOURCE, lineSettings);
    setSparkline(store, at(2, 4), 'C2:C5', lineSettings);
    expect(groups(wb)).toHaveLength(2);

    groupSparklines(store, [at(1, 4), at(2, 4)]);
    expect(groups(wb)).toHaveLength(1);
    expect(groups(wb)[0]?.sparklines.map((s) => s.location)).toEqual(['E2', 'E3']);

    ungroupSparklines(store, [at(1, 4)]);
    expect(groups(wb)).toHaveLength(2);
  });

  it('re-reads the group from the engine when it refuses a member change', async () => {
    const { wb, store } = await session();
    setSparkline(store, at(1, 4), SOURCE, lineSettings);
    const group = groupOfCell(store.getState(), at(1, 4));
    const members = vi.spyOn(wb, 'setSparklineGroupSparklines').mockReturnValue(false);
    mutators.setSparkline(store, at(2, 4), { source: 'C2:C5', group: group?.id as string });
    members.mockRestore();

    expect(listSparklines(store.getState()).map((entry) => entry.addr.row)).toEqual([1]);
    expect(groups(wb)[0]?.sparklines).toHaveLength(1);
  });
});
