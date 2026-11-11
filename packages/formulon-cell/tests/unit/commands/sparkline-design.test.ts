import { SparklineAxisType, SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  listSparklines,
  rgbColorSpec,
  setSparklines,
} from '../../../src/commands/sparkline.js';
import {
  clearSelectedSparklineGroupsAction,
  clearSelectedSparklinesAction,
  groupSelectedSparklinesAction,
  patchSelectedSparklineGroupsAction,
  replaceSparklineGroupAction,
  selectedSparklineCells,
  setSparklineAxisBoundAction,
  setSparklineColorAction,
  setSparklineSourceAction,
  switchSparklineRowColumnAction,
  toggleSparklineFlagAction,
  ungroupSelectedSparklinesAction,
} from '../../../src/commands/sparkline-design.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const at = (row: number, col = 5) => ({ sheet: 0, row, col });

const setup = () => {
  const store = createSpreadsheetStore();
  const history = new History();
  setSparklines(
    store,
    [0, 1, 2].map((row) => ({ addr: at(row), source: `A${row + 1}:C${row + 1}` })),
    defaultSparklineGroupSettings(SparklineType.Line),
    history,
  );
  mutators.setActive(store, at(0));
  mutators.setRange(store, { sheet: 0, r0: 0, c0: 5, r1: 2, c1: 5 });
  const ctx = (commandId: string) => ({ store, history, origin: 'ribbon' as const, commandId });
  const settings = (row = 0) => groupOfCell(store.getState(), at(row))?.settings;
  return { store, history, ctx, settings };
};

describe('sparkline design actions', () => {
  it('finds the sparklines inside the selection', () => {
    const { store } = setup();
    expect(selectedSparklineCells(store.getState()).map((a) => a.row)).toEqual([0, 1, 2]);
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 });
    expect(selectedSparklineCells(store.getState()).map((a) => a.row)).toEqual([1]);
  });

  it('sets the type of the group as one undo step', () => {
    const { ctx, settings, history } = setup();
    expect(
      patchSelectedSparklineGroupsAction(ctx('sparklineTypeColumn'), {
        type: SparklineType.Column,
      }),
    ).toBe(true);
    expect(settings()?.type).toBe(SparklineType.Column);
    expect(history.undo()).toBe(true);
    expect(settings()?.type).toBe(SparklineType.Line);
  });

  it('toggles a flag from the active group and undoes it in one step', () => {
    const { ctx, settings, history } = setup();
    toggleSparklineFlagAction(ctx('sparklineShowHigh'), 'high');
    expect(settings(1)?.high).toBe(true);
    toggleSparklineFlagAction(ctx('sparklineShowHigh'), 'high');
    expect(settings(1)?.high).toBe(false);
    history.undo();
    expect(settings(1)?.high).toBe(true);
  });

  it('sets series and marker colours', () => {
    const { ctx, settings } = setup();
    setSparklineColorAction(ctx('sparklineColor'), 'series', '#ff0000');
    setSparklineColorAction(ctx('sparklineMarkerColor'), 'high', '#00ff00');
    expect(settings()?.colorSeries).toEqual(rgbColorSpec('#ff0000'));
    expect(settings()?.colorHigh).toEqual(rgbColorSpec('#00ff00'));
  });

  it('sets the weight and the axis bounds', () => {
    const { ctx, settings } = setup();
    patchSelectedSparklineGroupsAction(ctx('sparklineColor'), { lineWidthPt: 2.25 });
    setSparklineAxisBoundAction(ctx('sparklineAxis'), 'min', 'custom', 5);
    setSparklineAxisBoundAction(ctx('sparklineAxis'), 'max', 'group');
    expect(settings()).toMatchObject({
      lineWidthPt: 2.25,
      minAxisType: SparklineAxisType.Custom,
      manualMin: 5,
      maxAxisType: SparklineAxisType.Group,
    });
  });

  it('ungroups into single groups and groups them back with the active settings', () => {
    const { store, ctx, settings, history } = setup();
    ungroupSelectedSparklinesAction(ctx('sparklineUngroup'));
    expect(store.getState().sparkline.groups.size).toBe(3);
    history.undo();
    expect(store.getState().sparkline.groups.size).toBe(1);
    ungroupSelectedSparklinesAction(ctx('sparklineUngroup'));
    patchSelectedSparklineGroupsAction(ctx('sparklineColor'), { lineWidthPt: 3 });
    mutators.setActive(store, at(1));
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 5, r1: 2, c1: 5 });
    groupSelectedSparklinesAction(ctx('sparklineGroup'));
    expect(store.getState().sparkline.groups.size).toBe(1);
    expect(settings(2)?.lineWidthPt).toBe(3);
  });

  it('clears the selected sparklines or their whole groups', () => {
    const { store, ctx } = setup();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 5, r1: 0, c1: 5 });
    expect(clearSelectedSparklinesAction(ctx('sparklineClear'))).toBe(1);
    expect(listSparklines(store.getState())).toHaveLength(2);
    mutators.setActive(store, at(1));
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 5, r1: 1, c1: 5 });
    expect(clearSelectedSparklineGroupsAction(ctx('sparklineClear'))).toBe(2);
    expect(listSparklines(store.getState())).toEqual([]);
  });

  it('replaces the members of a group, keeping its settings', () => {
    const { store, ctx, settings, history } = setup();
    patchSelectedSparklineGroupsAction(ctx('sparklineColor'), { lineWidthPt: 3 });
    const id = groupOfCell(store.getState(), at(0))?.id as string;
    replaceSparklineGroupAction(ctx('sparklineEditData'), id, [
      { addr: at(0), source: 'D1:F1' },
      { addr: at(4), source: 'D2:F2' },
    ]);
    expect(listSparklines(store.getState()).map(({ addr }) => addr.row)).toEqual([0, 4]);
    expect(settings(4)?.lineWidthPt).toBe(3);
    history.undo();
    expect(listSparklines(store.getState()).map(({ addr }) => addr.row)).toEqual([0, 1, 2]);
  });

  it('points one sparkline at other data without leaving its group', () => {
    const { store, ctx } = setup();
    setSparklineSourceAction(ctx('sparklineEditData'), at(1), 'D2:F2');
    const entries = listSparklines(store.getState());
    expect(entries[1]?.spec.source).toBe('D2:F2');
    expect(entries[1]?.spec.group).toBe(entries[0]?.spec.group);
  });

  it('switches row and column for a square block, and does nothing for others', () => {
    const { store, ctx, history } = setup();
    expect(switchSparklineRowColumnAction(ctx('sparklineEditData'))).toBe(true);
    expect(listSparklines(store.getState()).map(({ spec }) => spec.source)).toEqual([
      'A1:A3',
      'B1:B3',
      'C1:C3',
    ]);
    history.undo();
    setSparklineSourceAction(ctx('sparklineEditData'), at(2), 'A9:C9');
    expect(switchSparklineRowColumnAction(ctx('sparklineEditData'))).toBe(false);
  });

  it('refuses every change on a protected sheet', () => {
    const { store, ctx, settings } = setup();
    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    expect(
      patchSelectedSparklineGroupsAction(ctx('sparklineTypeColumn'), {
        type: SparklineType.Column,
      }),
    ).toBe(false);
    expect(ungroupSelectedSparklinesAction(ctx('sparklineUngroup'))).toBe(false);
    expect(clearSelectedSparklinesAction(ctx('sparklineClear'))).toBe(0);
    expect(settings()?.type).toBe(SparklineType.Line);
  });
});
