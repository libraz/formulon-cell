import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  applySessionChartColors,
  applySessionChartElement,
  chartElementState,
  chartUsesPalette,
  selectedSessionChart,
  switchSessionChartRowColumn,
} from '../../../src/commands/chart-design.js';
import { History } from '../../../src/commands/history.js';
import { createSessionChart, sessionChartById } from '../../../src/commands/session-chart.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SessionChart,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });

describe('chart design commands', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'North');
    wb.setText({ sheet: 0, row: 0, col: 2 }, 'South');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'Jan');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 10);
    wb.setNumber({ sheet: 0, row: 1, col: 2 }, 4);
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'Feb');
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 15);
    wb.setNumber({ sheet: 0, row: 2, col: 2 }, 6);
    wb.setText({ sheet: 0, row: 3, col: 0 }, 'Mar');
    wb.setNumber({ sheet: 0, row: 3, col: 1 }, 12);
    wb.setNumber({ sheet: 0, row: 3, col: 2 }, 9);
    store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    history = new History();
  });

  afterEach(() => wb.dispose());

  const chart = (kind: 'column' | 'bar' | 'pie' = 'column'): SessionChart => {
    const created = createSessionChart(store, range(0, 0, 3, kind === 'pie' ? 1 : 2), {
      wb,
      kind,
    });
    if (!created) throw new Error('chart was not created');
    return created;
  };
  const current = (id: string): SessionChart => {
    const found = sessionChartById(store.getState(), id);
    if (!found) throw new Error('chart is gone');
    return found;
  };

  it('finds the chart selected on the active sheet', () => {
    const created = chart();
    expect(selectedSessionChart(store.getState())).toBeNull();
    mutators.setSelectedObject(store, { id: created.id, name: 'Chart 1' });
    expect(selectedSessionChart(store.getState())?.id).toBe(created.id);
  });

  it('adds and removes axis titles, the chart title and data labels as one undo step each', () => {
    const { id } = chart();
    expect(chartElementState(current(id), 'axis-title-vertical')).toEqual({
      available: true,
      checked: false,
    });
    expect(applySessionChartElement(store, id, 'axis-title-vertical', history, wb)).toBe(true);
    expect(current(id).yAxis?.title).toEqual({ auto: true });
    expect(chartElementState(current(id), 'axis-title-vertical').checked).toBe(true);

    expect(chartElementState(current(id), 'title-above').checked).toBe(true);
    expect(applySessionChartElement(store, id, 'title-none', history, wb)).toBe(true);
    expect(current(id).title).toBe('');
    // Choosing what the chart already shows changes nothing.
    expect(applySessionChartElement(store, id, 'title-none', history, wb)).toBe(false);

    expect(applySessionChartElement(store, id, 'labels-show', history, wb)).toBe(true);
    expect(chartElementState(current(id), 'labels-show').checked).toBe(true);
    expect(applySessionChartElement(store, id, 'labels-none', history, wb)).toBe(true);
    expect(chartElementState(current(id), 'labels-none').checked).toBe(true);

    expect(history.undo()).toBe(true);
    expect(chartElementState(current(id), 'labels-show').checked).toBe(true);
    expect(history.undo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(current(id).title).toEqual({ auto: true });
    expect(history.undo()).toBe(true);
    expect(current(id).yAxis?.title).toBeUndefined();
  });

  it('moves and hides the legend', () => {
    const { id } = chart();
    expect(chartElementState(current(id), 'legend-bottom').checked).toBe(true);
    expect(applySessionChartElement(store, id, 'legend-right', history, wb)).toBe(true);
    expect(current(id).legend).toMatchObject({ position: 'right' });
    expect(chartElementState(current(id), 'legend-right').checked).toBe(true);
    expect(applySessionChartElement(store, id, 'legend-none', history, wb)).toBe(true);
    expect(chartElementState(current(id), 'legend-none').checked).toBe(true);
    expect(chartElementState(current(id), 'legend-right').checked).toBe(false);
  });

  it('offers gridlines only on the value axis, named by the direction they run', () => {
    const column = chart('column');
    expect(chartElementState(column, 'gridlines-horizontal')).toEqual({
      available: true,
      checked: true,
    });
    expect(chartElementState(column, 'gridlines-vertical').available).toBe(false);
    expect(applySessionChartElement(store, column.id, 'gridlines-horizontal', history, wb)).toBe(
      true,
    );
    expect(current(column.id).yAxis?.majorGridlines).toBe(false);

    const bar = chart('bar');
    expect(chartElementState(bar, 'gridlines-vertical').available).toBe(true);
    expect(chartElementState(bar, 'gridlines-horizontal').available).toBe(false);

    const pie = chart('pie');
    expect(chartElementState(pie, 'axis-title-horizontal').available).toBe(false);
    expect(applySessionChartElement(store, pie.id, 'axis-title-horizontal', history, wb)).toBe(
      false,
    );
  });

  it('switches rows and columns and back', () => {
    const { id } = chart();
    expect(current(id).data.series.map((series) => series.name)).toEqual(['North', 'South']);
    expect(switchSessionChartRowColumn(store, id, history, wb)).toBe(true);
    expect(current(id).data.series.map((series) => series.name)).toEqual(['Jan', 'Feb', 'Mar']);
    expect(switchSessionChartRowColumn(store, id, history, wb)).toBe(true);
    expect(current(id).data.series.map((series) => series.name)).toEqual(['North', 'South']);
    expect(history.undo()).toBe(true);
    expect(current(id).data.series).toHaveLength(3);
  });

  it('recolours the series with a preset palette', () => {
    const { id } = chart();
    expect(chartUsesPalette(current(id), 'office', wb)).toBe(true);
    expect(applySessionChartColors(store, id, 'mono', history, wb)).toBe(true);
    expect(chartUsesPalette(current(id), 'mono', wb)).toBe(true);
    expect(chartUsesPalette(current(id), 'office', wb)).toBe(false);
    expect(history.undo()).toBe(true);
    expect(chartUsesPalette(current(id), 'office', wb)).toBe(true);
  });
});
