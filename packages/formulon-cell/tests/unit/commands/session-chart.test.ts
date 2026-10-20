// @vitest-environment node
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  addSessionChartSeries,
  arrangeSessionChart,
  clearSessionChart,
  clearSessionChartsInRange,
  createRibbonChartFromSelection,
  createSessionChart,
  duplicateSessionChart,
  inferRecommendedChartKind,
  listSessionCharts,
  removeSessionChartSeries,
  sessionChartById,
  sessionChartsForRange,
  updateSessionChart,
} from '../../../src/commands/session-chart.js';
import {
  hydrateChartsFromEngine,
  refreshChartsAfterRecalc,
} from '../../../src/engine/charts-sync.js';
import { sessionPosition } from '../../../src/engine/drawing-geometry.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, '../../fixtures/charts', name)));

/** Two labelled series over three categories in A1:C4, a second table in E1:F3. */
const seed = (wb: WorkbookHandle): void => {
  const rows: (string | number)[][] = [
    ['', 'North', 'South'],
    ['Jan', 10, 20],
    ['Feb', 15, 25],
    ['Mar', 12, 22],
  ];
  rows.forEach((cells, row) => {
    cells.forEach((value, col) => {
      if (typeof value === 'number') wb.setNumber({ sheet: 0, row, col }, value);
      else if (value !== '') wb.setText({ sheet: 0, row, col }, value);
    });
  });
  wb.setText({ sheet: 0, row: 0, col: 4 }, 'k');
  wb.setText({ sheet: 0, row: 0, col: 5 }, 'v');
  wb.setText({ sheet: 0, row: 1, col: 4 }, 'x');
  wb.setNumber({ sheet: 0, row: 1, col: 5 }, 1);
  wb.setText({ sheet: 0, row: 2, col: 4 }, 'y');
  wb.setNumber({ sheet: 0, row: 2, col: 5 }, 2);
};

const engineChartIds = (wb: WorkbookHandle): number[] =>
  wb
    .listDrawingObjects(0)
    .filter((object) => object.kind === 2)
    .map((object) => object.objectId);

describe('session chart commands', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    seed(wb);
    store = createSpreadsheetStore();
    history = new History();
  });

  afterEach(() => {
    wb.dispose();
  });

  it('creates the chart in the workbook and projects what the engine reports', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
    if (!chart) throw new Error('chart was not created');

    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId]);
    expect(chart).toMatchObject({
      id: `chart-native-0-${chart.nativeObjectId}`,
      sheet: 0,
      kind: 'column',
      source: range(0, 0, 3, 2),
      seriesDirection: 'columns',
      legend: { visible: true, position: 'bottom' },
      nativeAnchor: { anchorKind: AnchorKind.TwoCell, editAs: AnchorEditAs.TwoCell },
    });
    // An omitted title is Excel's automatic one, which the renderer localizes.
    expect(chart.title).toEqual({ auto: true });
    expect(chart.data).toMatchObject({
      domain: 'category',
      categories: ['Jan', 'Feb', 'Mar'],
      series: [
        { id: 'series-0', name: 'North', values: [10, 15, 12] },
        { id: 'series-1', name: 'South', values: [20, 25, 22] },
      ],
    });
    expect(listSessionCharts(store.getState())).toEqual([chart]);
  });

  it('brings a chart to the front or sends it back, as one undo step', () => {
    const charts = [0, 1, 2].map((index) =>
      createSessionChart(store, range(0, 0, 3, 2), { wb, anchor: { row: index * 4, col: 6 } }),
    );
    const [first, second, third] = charts;
    if (!first || !second || !third) throw new Error('charts were not created');
    const order = (): number[] => listSessionCharts(store.getState()).map((c) => c.nativeObjectId);
    expect(order()).toEqual([first, second, third].map((c) => c.nativeObjectId));

    expect(arrangeSessionChart(store, first.id, 'bring-front', history, wb)).not.toBeNull();
    expect(order()).toEqual([second, third, first].map((c) => c.nativeObjectId));
    expect(engineChartIds(wb)).toEqual(order());
    expect(arrangeSessionChart(store, first.id, 'bring-front', history, wb)).toBeNull();

    expect(arrangeSessionChart(store, third.id, 'send-back', history, wb)).not.toBeNull();
    expect(order()).toEqual([third, second, first].map((c) => c.nativeObjectId));

    expect(history.undo()).toBe(true);
    expect(order()).toEqual([second, third, first].map((c) => c.nativeObjectId));
    expect(history.undo()).toBe(true);
    expect(order()).toEqual([first, second, third].map((c) => c.nativeObjectId));
    expect(history.redo()).toBe(true);
    expect(engineChartIds(wb)).toEqual([second, third, first].map((c) => c.nativeObjectId));

    setProtectedSheet(store, 0, true);
    expect(arrangeSessionChart(store, second.id, 'bring-front', history, wb)).toBeNull();
  });

  it('duplicates a chart on top of the original, offset from it, as one undo step', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), {
      wb,
      kind: 'line',
      title: 'Sales',
      anchor: { row: 2, col: 6 },
    });
    if (!chart) throw new Error('chart was not created');
    const at = (id: string) => {
      const found = sessionChartById(store.getState(), id);
      if (!found) throw new Error(`missing chart ${id}`);
      return sessionPosition(store, found.nativeAnchor);
    };

    const copy = duplicateSessionChart(store, chart.id, history, wb);
    if (!copy) throw new Error('chart was not duplicated');

    expect(copy.nativeObjectId).not.toBe(chart.nativeObjectId);
    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId, copy.nativeObjectId]);
    expect(copy).toMatchObject({
      kind: 'line',
      title: 'Sales',
      source: chart.source,
      legend: chart.legend,
    });
    expect(copy.data).toEqual(chart.data);
    const original = at(chart.id);
    const duplicated = at(copy.id);
    expect(duplicated.x).toBeCloseTo(original.x + 24, 0);
    expect(duplicated.y).toBeCloseTo(original.y + 24, 0);
    expect(duplicated.w).toBeCloseTo(original.w, 0);
    expect(duplicated.h).toBeCloseTo(original.h, 0);

    expect(history.undo()).toBe(true);
    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId]);
    expect(listSessionCharts(store.getState()).map((c) => c.id)).toEqual([chart.id]);
    expect(history.redo()).toBe(true);
    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId, copy.nativeObjectId]);
    expect(at(copy.id).x).toBeCloseTo(original.x + 24, 0);

    setProtectedSheet(store, 0, true);
    expect(duplicateSessionChart(store, chart.id, history, wb)).toBeNull();
    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId, copy.nativeObjectId]);
  });

  it('centres Excel’s default frame in the visible grid unless an anchor is given', () => {
    const centred = createSessionChart(store, range(0, 0, 3, 2), { wb });
    // 16 x 40 default cells of 64 x 20 px; the 480 x 288 px frame starts at (272, 256).
    expect(centred?.nativeAnchor).toMatchObject({
      fromCol: 4,
      fromColOff: 16 * 9525,
      fromRow: 12,
      fromRowOff: 16 * 9525,
      cx: 4_572_000,
      cy: 2_743_200,
    });

    const placed = createSessionChart(store, range(0, 0, 3, 2), {
      wb,
      anchor: { row: 20, col: 8, widthEmu: 952_500, heightEmu: 476_250 },
    });
    expect(placed?.nativeAnchor).toMatchObject({
      anchorKind: AnchorKind.TwoCell,
      fromRow: 20,
      fromCol: 8,
      cx: 952_500,
      cy: 476_250,
    });
  });

  it('writes the settings it is given and keeps Excel’s template for the rest', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), {
      wb,
      kind: 'bar',
      grouping: 'stacked',
      title: 'Sales',
      gapWidth: 80,
      overlap: 100,
      legend: { position: 'top' },
      palette: ['#112233', '#445566'],
      yAxis: { min: 0, max: 50 },
    });

    expect(chart).toMatchObject({
      kind: 'bar',
      grouping: 'stacked',
      title: 'Sales',
      gapWidth: 80,
      overlap: 100,
      legend: { visible: true, position: 'top' },
      seriesStyles: [
        { seriesId: 'series-0', fill: { color: '#112233' } },
        { seriesId: 'series-1', fill: { color: '#445566' } },
      ],
      // The template's gridlines stay where the options say nothing.
      xAxis: { visible: true, majorGridlines: false },
      yAxis: { visible: true, min: 0, max: 50, majorGridlines: true },
    });
    expect(chart?.palette).toBeUndefined();
    const info = wb.getChart(0, chart?.nativeObjectId ?? 0);
    expect(info?.title).toMatchObject({ text: 'Sales' });
  });

  it('stores a line series colour in its line and markers over the template’s', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), {
      wb,
      kind: 'line',
      palette: ['#112233', '#445566'],
      seriesStyles: [
        { seriesId: 'series-1', fill: { color: '#778899' }, marker: { visible: true } },
      ],
    });
    expect(chart?.seriesStyles).toMatchObject([
      { seriesId: 'series-0', line: { color: '#112233' }, marker: { visible: false } },
      {
        seriesId: 'series-1',
        line: { color: '#778899' },
        marker: { visible: true, fill: { color: '#778899' }, line: { color: '#778899' } },
      },
    ]);
    expect(chart?.seriesStyles?.every((style) => style.fill === undefined)).toBe(true);

    // A new palette replaces the colours the line and markers store.
    const updated = updateSessionChart(
      store,
      chart?.id ?? '',
      { palette: ['#aa0000', '#00aa00'] },
      new History(),
      wb,
    );
    expect(updated?.seriesStyles).toMatchObject([
      { seriesId: 'series-0', line: { color: '#aa0000' } },
      {
        seriesId: 'series-1',
        line: { color: '#00aa00' },
        marker: { fill: { color: '#00aa00' }, line: { color: '#00aa00' } },
      },
    ]);
  });

  it('creates nothing when the engine has no chart support', async () => {
    const stub = await WorkbookHandle.createDefault({ preferStub: true });
    try {
      expect(stub.capabilities.charts).not.toBe(true);
      expect(createSessionChart(store, range(0, 0, 3, 2), { wb: stub })).toBeNull();
      expect(store.getState().charts.charts).toEqual([]);
    } finally {
      stub.dispose();
    }
  });

  it('undoes a creation by removing the chart and redoes it under the same id', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb, kind: 'line' }, history);
    if (!chart) throw new Error('chart was not created');

    expect(history.undo()).toBe(true);
    expect(engineChartIds(wb)).toEqual([]);
    expect(store.getState().charts.charts).toEqual([]);

    expect(history.redo()).toBe(true);
    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId]);
    expect(sessionChartById(store.getState(), chart.id)).toEqual(chart);
  });

  it('applies a settings change through the engine and undoes it from a snapshot', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb, title: 'Before' });
    if (!chart) throw new Error('chart was not created');

    const updated = updateSessionChart(
      store,
      chart.id,
      { title: 'After', gapWidth: 40, dataLabels: { visible: true, showValue: true } },
      history,
      wb,
    );
    expect(updated).toMatchObject({
      title: 'After',
      gapWidth: 40,
      nativeObjectId: chart.nativeObjectId,
    });
    expect(updated?.dataLabels).toMatchObject({ visible: true, showValue: true });
    expect(wb.getChart(0, chart.nativeObjectId)?.title).toMatchObject({ text: 'After' });

    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toEqual(chart);
    expect(wb.getChart(0, chart.nativeObjectId)?.title).toMatchObject({ text: 'Before' });

    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toEqual(updated);
  });

  it('re-reads the chart after a type change and writes the rest against the new plot', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
    if (!chart) throw new Error('chart was not created');

    const line = updateSessionChart(
      store,
      chart.id,
      {
        kind: 'line',
        seriesStyles: [{ seriesId: 'series-1', line: { width: 3, dash: 'dash' } }],
      },
      history,
      wb,
    );
    expect(line).toMatchObject({ kind: 'line', grouping: 'clustered' });
    expect(line?.seriesStyles?.find((style) => style.seriesId === 'series-1')?.line).toMatchObject({
      width: 3,
      dash: 'dash',
    });
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)?.kind).toBe('column');
  });

  it('re-splits the series when the data range changes', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
    if (!chart) throw new Error('chart was not created');

    const moved = updateSessionChart(store, chart.id, { source: range(0, 4, 2, 5) }, history, wb);
    expect(moved?.source).toEqual(range(0, 4, 2, 5));
    expect(moved?.data).toMatchObject({
      categories: ['x', 'y'],
      series: [{ name: 'v', values: [1, 2] }],
    });
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)?.source).toEqual(range(0, 0, 3, 2));
  });

  it('switches the series direction through the engine split', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
    if (!chart) throw new Error('chart was not created');
    const rows = updateSessionChart(store, chart.id, { seriesDirection: 'rows' }, history, wb);
    expect(rows?.seriesDirection).toBe('rows');
    expect(rows?.data).toMatchObject({
      categories: ['North', 'South'],
      series: [{ name: 'Jan' }, { name: 'Feb' }, { name: 'Mar' }],
    });
  });

  describe('series add and remove', () => {
    /** A third series beside the A1:C4 table: West in D1:D4. */
    const seedWest = (): void => {
      wb.setText({ sheet: 0, row: 0, col: 3 }, 'West');
      [30, 35, 32].forEach((value, i) => {
        wb.setNumber({ sheet: 0, row: i + 1, col: 3 }, value);
      });
    };

    it('adds a series from cell references and undoes and redoes it', () => {
      seedWest();
      const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
      if (!chart) throw new Error('chart was not created');

      const added = addSessionChartSeries(
        store,
        chart.id,
        { name: range(0, 3, 0, 3), values: range(1, 3, 3, 3), categories: range(1, 0, 3, 0) },
        history,
        wb,
      );
      expect(added?.data).toMatchObject({
        domain: 'category',
        categories: ['Jan', 'Feb', 'Mar'],
        series: [
          { id: 'series-0', name: 'North' },
          { id: 'series-1', name: 'South' },
          { id: 'series-2', name: 'West', values: [30, 35, 32] },
        ],
      });
      expect(wb.getChart(0, chart.nativeObjectId)?.plots[0]?.series[2]?.values).toMatchObject({
        formula: 'Sheet1!$D$2:$D$4',
      });
      expect(wb.getChartData(0, chart.nativeObjectId)?.series).toHaveLength(3);

      expect(history.undo()).toBe(true);
      expect(sessionChartById(store.getState(), chart.id)).toEqual(chart);
      expect(wb.getChartData(0, chart.nativeObjectId)?.series).toHaveLength(2);

      expect(history.redo()).toBe(true);
      expect(sessionChartById(store.getState(), chart.id)).toEqual(added);
    });

    it('adds a scatter series from X and Y values', () => {
      seedWest();
      const chart = createSessionChart(store, range(0, 1, 3, 2), { wb, kind: 'scatter' });
      if (!chart) throw new Error('chart was not created');
      const before = chart.data.series.length;

      const added = addSessionChartSeries(
        store,
        chart.id,
        { name: range(0, 3, 0, 3), values: range(1, 3, 3, 3), x: range(1, 1, 3, 1) },
        history,
        wb,
      );
      expect(added?.data.domain).toBe('xy');
      expect(added?.data.series).toHaveLength(before + 1);
      const series = added?.data.series[before];
      expect(series?.name).toBe('West');
      expect(series && 'points' in series ? series.points.map(({ x, y }) => [x, y]) : []).toEqual([
        [10, 30],
        [15, 35],
        [12, 32],
      ]);
    });

    it('removes a series, keeps the last one, and undoes the removal', () => {
      const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
      if (!chart) throw new Error('chart was not created');

      const removed = removeSessionChartSeries(store, chart.id, 'series-0', history, wb);
      expect(removed?.data.series.map((series) => series.name)).toEqual(['South']);
      expect(wb.getChart(0, chart.nativeObjectId)?.plots[0]?.series).toHaveLength(1);

      // The last series stays.
      expect(removeSessionChartSeries(store, chart.id, 'series-1', history, wb)).toBeNull();
      expect(removeSessionChartSeries(store, chart.id, 'series-9', history, wb)).toBeNull();
      expect(wb.getChart(0, chart.nativeObjectId)?.plots[0]?.series).toHaveLength(1);

      expect(history.undo()).toBe(true);
      expect(sessionChartById(store.getState(), chart.id)).toEqual(chart);
      expect(history.redo()).toBe(true);
      expect(sessionChartById(store.getState(), chart.id)).toEqual(removed);
    });

    it('rejects series edits on a protected sheet and without chart support', async () => {
      seedWest();
      const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
      if (!chart) throw new Error('chart was not created');
      const input = { values: range(0, 3, 1, 3) };
      setProtectedSheet(store, 0, true);
      expect(addSessionChartSeries(store, chart.id, input, history, wb)).toBeNull();
      expect(removeSessionChartSeries(store, chart.id, 'series-0', history, wb)).toBeNull();
      expect(wb.getChart(0, chart.nativeObjectId)?.plots[0]?.series).toHaveLength(2);
      expect(history.canUndo()).toBe(false);
      setProtectedSheet(store, 0, false);

      const stub = await WorkbookHandle.createDefault({ preferStub: true });
      try {
        expect(addSessionChartSeries(store, chart.id, input, history, stub)).toBeNull();
        expect(removeSessionChartSeries(store, chart.id, 'series-0', history, stub)).toBeNull();
      } finally {
        stub.dispose();
      }
      expect(listSessionCharts(store.getState())).toEqual([chart]);
    });

    it('changes nothing when the engine refuses a series', async () => {
      const loaded = await WorkbookHandle.loadBytes(fixture('chart_pivot.xlsx'));
      try {
        hydrateChartsFromEngine(loaded, store, 1);
        const [chart] = listSessionCharts(store.getState());
        if (!chart) throw new Error('fixture chart missing');
        const before = loaded.snapshotDrawing(1, chart.nativeObjectId);
        // A PivotChart's series follow its PivotTable.
        expect(
          addSessionChartSeries(
            store,
            chart.id,
            { values: { sheet: 1, r0: 0, c0: 0, r1: 2, c1: 0 } },
            history,
            loaded,
          ),
        ).toBeNull();
        expect(loaded.snapshotDrawing(1, chart.nativeObjectId)).toEqual(before);
        expect(sessionChartById(store.getState(), chart.id)).toEqual(chart);
        expect(history.canUndo()).toBe(false);
      } finally {
        loaded.dispose();
      }
    });
  });

  it('moves a chart by re-anchoring it, keeping its id and two-cell anchor', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), {
      wb,
      anchor: { row: 5, col: 5, widthEmu: 1_905_000, heightEmu: 952_500 },
    });
    if (!chart) throw new Error('chart was not created');

    const moved = updateSessionChart(
      store,
      chart.id,
      { anchor: { row: 9, col: 2, widthEmu: 2_857_500, heightEmu: 952_500 } },
      history,
      wb,
    );
    expect(moved?.nativeObjectId).toBe(chart.nativeObjectId);
    expect(moved?.nativeAnchor).toMatchObject({
      anchorKind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 9,
      fromCol: 2,
      cx: 2_857_500,
    });

    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)?.nativeAnchor).toMatchObject({
      anchorKind: AnchorKind.TwoCell,
      fromRow: 5,
      fromCol: 5,
      cx: 1_905_000,
    });
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)?.nativeAnchor).toMatchObject({
      fromRow: 9,
      fromCol: 2,
    });
  });

  it('deletes a chart and restores it under the same id', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
    if (!chart) throw new Error('chart was not created');

    expect(clearSessionChart(store, chart.id, history, wb)).toBe(true);
    expect(engineChartIds(wb)).toEqual([]);
    expect(listSessionCharts(store.getState())).toEqual([]);

    expect(history.undo()).toBe(true);
    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId]);
    expect(sessionChartById(store.getState(), chart.id)).toEqual(chart);

    expect(history.redo()).toBe(true);
    expect(engineChartIds(wb)).toEqual([]);
  });

  it('deletes the charts of a range in one undo step', () => {
    const a = createSessionChart(store, range(0, 0, 3, 2), { wb });
    const b = createSessionChart(store, range(0, 4, 2, 5), { wb });
    if (!a || !b) throw new Error('charts were not created');

    expect(sessionChartsForRange(store.getState(), range(1, 1, 1, 1))).toEqual([a]);
    expect(clearSessionChartsInRange(store, range(1, 1, 1, 1), history, wb)).toBe(1);
    expect(listSessionCharts(store.getState()).map((chart) => chart.id)).toEqual([b.id]);
    expect(history.undo()).toBe(true);
    // Restored in its stacking slot, below the later chart.
    expect(listSessionCharts(store.getState()).map((chart) => chart.id)).toEqual([a.id, b.id]);
  });

  it('rejects create, update and delete on a protected sheet', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), { wb });
    if (!chart) throw new Error('chart was not created');
    setProtectedSheet(store, 0, true);

    expect(createSessionChart(store, range(0, 0, 3, 2), { wb })).toBeNull();
    expect(updateSessionChart(store, chart.id, { title: 'x' }, null, wb)).toBeNull();
    expect(updateSessionChart(store, chart.id, { anchor: { row: 9 } }, null, wb)).toBeNull();
    expect(clearSessionChart(store, chart.id, null, wb)).toBe(false);
    expect(clearSessionChartsInRange(store, range(0, 0, 3, 2), null, wb)).toBe(0);
    expect(engineChartIds(wb)).toEqual([chart.nativeObjectId]);
    expect(listSessionCharts(store.getState())).toEqual([chart]);
  });

  it('checks the policy on commands and on replay', () => {
    const controller = new InteractionController({ store, history, getWb: () => wb });
    const unregister = registerInteractionController(store, controller);
    try {
      const chart = createSessionChart(store, range(0, 0, 3, 2), { wb }, history);
      if (!chart) throw new Error('chart was not created');
      updateSessionChart(store, chart.id, { title: 'Edited' }, history, wb);
      expect(history.undo()).toBe(true);
      controller.setPolicy(viewerPolicy());
      expect(history.redo()).toBe(false);
      expect(history.undo()).toBe(false);
      expect(updateSessionChart(store, chart.id, { title: 'Denied' }, history, wb)).toBeNull();
      expect(clearSessionChart(store, chart.id, history, wb)).toBe(false);
      expect(createSessionChart(store, range(0, 0, 3, 2), { wb }, history)).toBeNull();
      expect(engineChartIds(wb)).toEqual([chart.nativeObjectId]);
      expect(sessionChartById(store.getState(), chart.id)?.title).toEqual({ auto: true });
    } finally {
      unregister();
      controller.dispose();
    }
  });

  it('hands out detached copies', () => {
    const chart = createSessionChart(store, range(0, 0, 3, 2), {
      wb,
      seriesStyles: [{ seriesId: 'series-0', dataLabels: { visible: true, showValue: true } }],
    });
    if (!chart) throw new Error('chart was not created');
    const style = chart.seriesStyles?.[0];
    expect(style?.dataLabels).toBeDefined();
    (style?.dataLabels as { showValue: boolean }).showValue = false;
    (chart.nativeAnchor as { fromRow: number }).fromRow = 999;
    const stored = sessionChartById(store.getState(), chart.id);
    expect(stored?.seriesStyles?.[0]?.dataLabels?.showValue).toBe(true);
    expect(stored?.nativeAnchor.fromRow).not.toBe(999);
  });

  it('re-reads a chart over formula cells when a recalc changes their inputs', () => {
    // H1:I3 shows v = 10 * the North values; the chart reads only the formulas.
    wb.setText({ sheet: 0, row: 0, col: 7 }, 'k');
    wb.setText({ sheet: 0, row: 0, col: 8 }, 'v');
    wb.setText({ sheet: 0, row: 1, col: 7 }, 'a');
    wb.setText({ sheet: 0, row: 2, col: 7 }, 'b');
    wb.setFormula({ sheet: 0, row: 1, col: 8 }, '=B2*10');
    wb.setFormula({ sheet: 0, row: 2, col: 8 }, '=B3*10');
    const chart = createSessionChart(store, range(0, 7, 2, 8), { wb });
    if (!chart) throw new Error('chart was not created');
    expect(chart.data).toMatchObject({ series: [{ values: [100, 150] }] });
    const unsubscribe = wb.subscribe((event) => {
      if (event.kind === 'recalc') refreshChartsAfterRecalc(wb, store);
    });
    try {
      wb.setNumber({ sheet: 0, row: 1, col: 1 }, 7);
    } finally {
      unsubscribe();
    }
    expect(sessionChartById(store.getState(), chart.id)?.data).toMatchObject({
      series: [{ values: [70, 150] }],
    });
  });

  it('projects workbook charts it cannot draw as placeholders that only move', async () => {
    const loaded = await WorkbookHandle.loadBytes(fixture('chart_chartex.xlsx'));
    try {
      hydrateChartsFromEngine(loaded, store, 0);
      const [chart] = listSessionCharts(store.getState());
      if (!chart) throw new Error('fixture chart missing');
      expect(chart).toMatchObject({ kind: 'unsupported', unsupported: { reason: 'chartex' } });
      expect(chart.source).toBeUndefined();
      expect(sessionChartsForRange(store.getState(), range(0, 0, 100, 100))).toEqual([]);
      expect(updateSessionChart(store, chart.id, { title: 'x' }, null, loaded)).toBeNull();
      const moved = updateSessionChart(
        store,
        chart.id,
        { anchor: { row: 30, col: 1 } },
        null,
        loaded,
      );
      expect(moved?.nativeAnchor).toMatchObject({ fromRow: 30, fromCol: 1 });
    } finally {
      loaded.dispose();
    }
  });

  it('restores the chart and changes nothing when the engine rejects a step', async () => {
    const loaded = await WorkbookHandle.loadBytes(fixture('chart_pivot.xlsx'));
    try {
      hydrateChartsFromEngine(loaded, store, 1);
      const [chart] = listSessionCharts(store.getState());
      if (!chart) throw new Error('fixture chart missing');
      expect(chart.pivotChart).toBe(true);
      const before = loaded.snapshotDrawing(1, chart.nativeObjectId);
      // A PivotChart's series follow the PivotTable; the engine refuses to re-split them.
      expect(
        updateSessionChart(
          store,
          chart.id,
          {
            title: 'Renamed',
            seriesDirection: chart.seriesDirection === 'rows' ? 'columns' : 'rows',
          },
          history,
          loaded,
        ),
      ).toBeNull();
      expect(loaded.snapshotDrawing(1, chart.nativeObjectId)).toEqual(before);
      expect(sessionChartById(store.getState(), chart.id)).toEqual(chart);
      expect(history.canUndo()).toBe(false);
    } finally {
      loaded.dispose();
    }
  });

  it('inserts ribbon charts with the recommended kind', () => {
    expect(inferRecommendedChartKind(range(0, 0, 0, 4))).toBe('line');
    expect(inferRecommendedChartKind(range(0, 0, 4, 0))).toBe('bar');
    expect(inferRecommendedChartKind(range(0, 0, 3, 1))).toBe('pie');
    expect(inferRecommendedChartKind(range(0, 0, 3, 2))).toBe('column');
    const chart = createRibbonChartFromSelection({
      store,
      wb,
      range: range(0, 4, 2, 5),
      action: 'recommended',
      history,
    });
    expect(chart?.kind).toBe('pie');
    expect(history.canUndo()).toBe(true);
  });
});
