import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  clearSessionChart,
  clearSessionChartsInRange,
  createSessionChart,
  listSessionCharts,
  sessionChartById,
  sessionChartData,
  sessionChartSeries,
  sessionChartsForRange,
  updateSessionChart,
} from '../../../src/commands/session-chart.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) =>
  ({ sheet: 0, r0, c0, r1, c1 }) as const;

const seedNumbers = (
  store: ReturnType<typeof createSpreadsheetStore>,
  entries: Array<{ row: number; col: number; value: number }>,
): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    for (const e of entries) {
      cells.set(addrKey({ sheet: 0, row: e.row, col: e.col }), {
        value: { kind: 'number', value: e.value },
        formula: null,
      });
    }
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('session chart commands', () => {
  it('preserves log bases and dormant linear options through type changes and history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const chart = createSessionChart(store, range(0, 0, 2, 1), {
      id: 'log-axes',
      kind: 'scatter',
      xAxis: { logBase: 2, reverseOrder: true, majorUnit: 5 },
      yAxis: { logBase: 10, min: 1, max: 1000, crossing: 10 },
    });
    if (!chart) throw new Error('Chart creation was rejected');
    expect(chart.xAxis).toEqual({ logBase: 2, reverseOrder: true, majorUnit: 5 });
    updateSessionChart(store, chart.id, { kind: 'doughnut' }, history);
    expect(sessionChartById(store.getState(), chart.id)?.yAxis?.logBase).toBe(10);
    updateSessionChart(
      store,
      chart.id,
      {
        kind: 'column',
        yAxis: { ...chart.yAxis, logBase: 1000 },
      },
      history,
    );
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toMatchObject({
      kind: 'doughnut',
      yAxis: { logBase: 10, min: 1, max: 1000, crossing: 10 },
    });
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toMatchObject({
      kind: 'column',
      xAxis: { logBase: 2, reverseOrder: true, majorUnit: 5 },
      yAxis: { logBase: 1000, min: 1, max: 1000, crossing: 10 },
    });
  });

  it('preserves axis label positions across creation, kind changes, and history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const chart = createSessionChart(store, range(0, 0, 2, 1), {
      id: 'positioned-labels',
      kind: 'scatter',
      xAxis: { labelPosition: 'high', crossing: 2, reverseOrder: true },
      yAxis: { labelPosition: 'nextTo', title: 'Value' },
    });
    if (!chart) throw new Error('Chart creation was rejected');
    expect(chart.xAxis).toEqual({ labelPosition: 'high', crossing: 2, reverseOrder: true });
    expect(chart.yAxis).toEqual({ labelPosition: 'nextTo', title: 'Value' });
    updateSessionChart(store, chart.id, { kind: 'doughnut' }, history);
    expect(sessionChartById(store.getState(), chart.id)?.xAxis?.labelPosition).toBe('high');
    updateSessionChart(
      store,
      chart.id,
      {
        kind: 'bar',
        yAxis: { ...chart.yAxis, labelPosition: 'none' },
      },
      history,
    );
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toMatchObject({
      kind: 'doughnut',
      yAxis: { labelPosition: 'nextTo', title: 'Value' },
    });
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toMatchObject({
      kind: 'bar',
      xAxis: { labelPosition: 'high', crossing: 2, reverseOrder: true },
      yAxis: { labelPosition: 'none', title: 'Value' },
    });
  });

  it('preserves axis crossings through type changes and undo/redo', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const chart = createSessionChart(store, range(0, 0, 2, 1), {
      id: 'crossed-axes',
      kind: 'scatter',
      xAxis: { crossing: 'max', reverseOrder: true, min: -10 },
      yAxis: { crossing: 2.5, numberFormat: '0.00' },
    });
    if (!chart) throw new Error('Chart creation was rejected');
    expect(chart.xAxis).toEqual({ crossing: 'max', reverseOrder: true, min: -10 });
    expect(chart.yAxis).toEqual({ crossing: 2.5, numberFormat: '0.00' });
    updateSessionChart(store, chart.id, { kind: 'doughnut' }, history);
    expect(sessionChartById(store.getState(), chart.id)?.yAxis?.crossing).toBe(2.5);
    updateSessionChart(
      store,
      chart.id,
      {
        kind: 'bar',
        xAxis: { ...chart.xAxis, crossing: 'min' },
        yAxis: { ...chart.yAxis, crossing: 2 },
      },
      history,
    );
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toMatchObject({
      kind: 'doughnut',
      xAxis: { crossing: 'max' },
      yAxis: { crossing: 2.5 },
    });
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)).toMatchObject({
      kind: 'bar',
      xAxis: { crossing: 'min', reverseOrder: true, min: -10 },
      yAxis: { crossing: 2, numberFormat: '0.00' },
    });
  });

  it('keeps physical-axis reversal options through creation and history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const chart = createSessionChart(store, range(0, 0, 2, 1), {
      id: 'reversed-axes',
      kind: 'scatter',
      xAxis: { reverseOrder: true, min: -10, max: 20 },
      yAxis: { reverseOrder: false, numberFormat: '0.00' },
    });
    expect(chart?.xAxis).toEqual({ reverseOrder: true, min: -10, max: 20 });
    expect(chart?.yAxis).toEqual({ reverseOrder: false, numberFormat: '0.00' });
    if (!chart) throw new Error('Chart creation was rejected');
    expect(
      updateSessionChart(
        store,
        chart.id,
        {
          xAxis: { ...chart.xAxis, reverseOrder: false },
          yAxis: { ...chart.yAxis, reverseOrder: true },
        },
        history,
      ),
    ).toMatchObject({
      xAxis: { reverseOrder: false, min: -10, max: 20 },
      yAxis: { reverseOrder: true, numberFormat: '0.00' },
    });
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)?.xAxis?.reverseOrder).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)?.yAxis?.reverseOrder).toBe(false);
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), chart.id)?.xAxis?.reverseOrder).toBe(false);
    expect(sessionChartById(store.getState(), chart.id)?.yAxis?.reverseOrder).toBe(true);
  });

  it('preserves scatter style across creation, type changes, and history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const chart = createSessionChart(store, range(0, 0, 2, 1), {
      id: 'connected-scatter',
      kind: 'scatter',
      scatterStyle: 'lineMarker',
    });
    expect(chart?.scatterStyle).toBe('lineMarker');
    expect(updateSessionChart(store, 'connected-scatter', { kind: 'column' }, history)?.kind).toBe(
      'column',
    );
    expect(sessionChartById(store.getState(), 'connected-scatter')?.scatterStyle).toBe(
      'lineMarker',
    );
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), 'connected-scatter')).toMatchObject({
      kind: 'scatter',
      scatterStyle: 'lineMarker',
    });
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), 'connected-scatter')?.kind).toBe('column');
    expect(
      updateSessionChart(
        store,
        'connected-scatter',
        { kind: 'scatter', scatterStyle: 'line' },
        history,
      )?.scatterStyle,
    ).toBe('line');
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), 'connected-scatter')?.scatterStyle).toBe(
      'lineMarker',
    );
    const implicit = createSessionChart(store, range(0, 0, 2, 1), {
      id: 'implicit-scatter',
      kind: 'scatter',
    });
    expect(implicit?.scatterStyle).toBeUndefined();
  });

  it('keeps doughnut geometry and dormant settings through type changes and history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    seedNumbers(store, [
      { row: 0, col: 0, value: 10 },
      { row: 0, col: 1, value: 20 },
      { row: 1, col: 0, value: 30 },
      { row: 1, col: 1, value: 40 },
    ]);
    const sourceCells = store.getState().data.cells;
    const chart = createSessionChart(store, range(0, 0, 1, 1), {
      id: 'rings',
      kind: 'doughnut',
      holeSize: 70,
      firstSliceAngle: 90,
      grouping: 'stacked',
      yAxis: { min: -10, max: 50, numberFormat: '0.00' },
      filter: { hiddenCategoryIds: ['category-1'] },
    });
    expect(chart?.title).toBe('Doughnut chart');
    expect(chart?.holeSize).toBe(70);
    if (!chart) throw new Error('Doughnut chart creation was rejected');
    const data = sessionChartData(store.getState(), chart);
    expect(data.domain).toBe('category');
    expect(data.series).toHaveLength(2);
    expect(updateSessionChart(store, 'rings', { kind: 'column' }, history)).toMatchObject({
      kind: 'column',
    });
    expect(sessionChartById(store.getState(), 'rings')).toMatchObject({
      kind: 'column',
      holeSize: 70,
      firstSliceAngle: 90,
      grouping: 'stacked',
      yAxis: { min: -10, max: 50, numberFormat: '0.00' },
      filter: { hiddenCategoryIds: ['category-1'] },
    });
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), 'rings')?.kind).toBe('doughnut');
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), 'rings')?.kind).toBe('column');
    expect(store.getState().data.cells).toBe(sourceCells);
  });

  it('reports the replay origin consistently to a host history guard', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    createSessionChart(store, range(0, 0, 2, 2), { id: 'origin' });
    updateSessionChart(store, 'origin', { dataLabels: { visible: true } }, history, {
      origin: 'contextMenu',
    });
    const origins: string[] = [];
    history.setGuard((entry, direction) => {
      const intent = direction === 'undo' ? entry.inverseIntent : entry.intent;
      origins.push(intent?.origin ?? 'missing');
      return intent?.origin === direction;
    });
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), 'origin')?.dataLabels).toBeUndefined();
    expect(history.redo()).toBe(true);
    expect(sessionChartById(store.getState(), 'origin')?.dataLabels?.visible).toBe(true);
    expect(origins).toEqual(['undo', 'redo']);
  });
  it('detaches nested settings from callers, patches, returned charts and history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const hiddenSeriesIds = ['series-1'];
    const markerFill = { color: '#123456' };
    const pointFill = { color: '#abcdef' };
    const styles = [
      {
        seriesId: 'series-0',
        marker: { fill: markerFill },
        points: [{ pointId: 'category-0', fill: pointFill }],
      },
    ];
    const chart = createSessionChart(
      store,
      range(0, 0, 2, 2),
      {
        id: 'nested',
        filter: { hiddenSeriesIds },
        dataLabels: { visible: true, color: '#112233' },
        seriesStyles: styles,
      },
      history,
    );
    hiddenSeriesIds.push('series-0');
    markerFill.color = '#000000';
    pointFill.color = '#000000';
    expect(sessionChartById(store.getState(), 'nested')?.filter?.hiddenSeriesIds).toEqual([
      'series-1',
    ]);
    expect(
      sessionChartById(store.getState(), 'nested')?.seriesStyles?.[0]?.marker?.fill?.color,
    ).toBe('#123456');
    const detached = chart?.seriesStyles?.[0]?.points?.[0]?.fill;
    if (!detached) throw new Error('missing detached point style');
    (detached as { color: string }).color = '#ffffff';
    const patchFill = { color: '#ff0000' };
    updateSessionChart(
      store,
      'nested',
      { seriesStyles: [{ seriesId: 'series-0', fill: patchFill }] },
      history,
    );
    patchFill.color = '#000000';
    expect(sessionChartById(store.getState(), 'nested')?.seriesStyles?.[0]?.fill?.color).toBe(
      '#ff0000',
    );
    history.undo();
    expect(
      sessionChartById(store.getState(), 'nested')?.seriesStyles?.[0]?.points?.[0]?.fill?.color,
    ).toBe('#abcdef');
    const list = listSessionCharts(store.getState());
    const listMarker = list[0]?.seriesStyles?.[0]?.marker?.fill;
    if (!listMarker) throw new Error('missing detached marker');
    (listMarker as { color: string }).color = '#000000';
    history.redo();
    expect(sessionChartById(store.getState(), 'nested')?.seriesStyles?.[0]?.fill?.color).toBe(
      '#ff0000',
    );
    history.undo();
    expect(
      sessionChartById(store.getState(), 'nested')?.seriesStyles?.[0]?.marker?.fill?.color,
    ).toBe('#123456');
  });

  it('rejects chart settings, insertion, deletion and replay in a viewer', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({ store, history, getWb: () => wb });
    const unregister = registerInteractionController(store, controller);
    try {
      const chart = createSessionChart(store, range(0, 0, 2, 2), { id: 'policy' }, history);
      expect(chart).not.toBeNull();
      updateSessionChart(store, 'policy', { dataLabels: { visible: true } }, history);
      expect(history.undo()).toBe(true);
      controller.setPolicy(viewerPolicy());
      expect(history.redo()).toBe(false);
      expect(history.undo()).toBe(false);
      expect(
        updateSessionChart(store, 'policy', { dataLabels: { visible: true } }, history),
      ).toBeNull();
      expect(clearSessionChart(store, 'policy', history)).toBe(false);
      expect(clearSessionChartsInRange(store, range(0, 0, 2, 2), history)).toBe(0);
      expect(createSessionChart(store, range(0, 0, 2, 2), { id: 'denied' }, history)).toBeNull();
      expect(store.getState().charts.charts).toHaveLength(1);
      expect(sessionChartById(store.getState(), 'policy')?.dataLabels).toBeUndefined();
    } finally {
      unregister();
      controller.dispose();
      wb.dispose();
    }
  });

  it('checks both protected sheets when a settings update moves its source', () => {
    const store = createSpreadsheetStore();
    createSessionChart(store, range(0, 0, 2, 2), { id: 'source' });
    setProtectedSheet(store, 1, true);
    expect(
      updateSessionChart(store, 'source', { source: { ...range(0, 0, 2, 2), sheet: 1 } }),
    ).toBeNull();
    expect(sessionChartById(store.getState(), 'source')?.source.sheet).toBe(0);
  });
  it('preserves Excel geometry options through create, update, undo and redo', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const chart = createSessionChart(
      store,
      range(0, 0, 2, 2),
      {
        id: 'geometry',
        gapWidth: 150,
        overlap: -25,
        firstSliceAngle: 90,
      },
      history,
    );
    expect(chart).toMatchObject({ gapWidth: 150, overlap: -25, firstSliceAngle: 90 });
    updateSessionChart(
      store,
      'geometry',
      { gapWidth: 0, overlap: 100, firstSliceAngle: 270 },
      history,
    );
    expect(sessionChartById(store.getState(), 'geometry')).toMatchObject({
      gapWidth: 0,
      overlap: 100,
      firstSliceAngle: 270,
    });
    history.undo();
    expect(sessionChartById(store.getState(), 'geometry')).toMatchObject({
      gapWidth: 150,
      overlap: -25,
      firstSliceAngle: 90,
    });
    history.redo();
    expect(sessionChartById(store.getState(), 'geometry')).toMatchObject({
      gapWidth: 0,
      overlap: 100,
      firstSliceAngle: 270,
    });
  });
  it('creates a reusable session chart overlay with stable defaults', () => {
    const store = createSpreadsheetStore();
    const r = range(1, 0, 5, 2);

    const chart = createSessionChart(store, r, { kind: 'line' });

    expect(chart).toEqual({
      id: 'chart-0-1-0-5-2-line',
      kind: 'line',
      source: r,
      title: 'Line chart',
    });
    expect(store.getState().charts.charts).toEqual([chart]);
  });

  it('supports Excel-style chart kinds used by the ribbon', () => {
    const store = createSpreadsheetStore();

    expect(createSessionChart(store, range(0, 0, 3, 1), { kind: 'bar' })?.title).toBe('Bar chart');
    expect(createSessionChart(store, range(0, 0, 3, 1), { kind: 'area' })?.title).toBe(
      'Area chart',
    );
    expect(createSessionChart(store, range(0, 0, 3, 1), { kind: 'pie' })?.title).toBe('Pie chart');
    expect(createSessionChart(store, range(0, 0, 3, 1), { kind: 'scatter' })?.title).toBe(
      'Scatter chart',
    );
  });

  it('can leave the title unset for localized render fallbacks', () => {
    const store = createSpreadsheetStore();
    const chart = createSessionChart(store, range(0, 0, 2, 0), {
      kind: 'column',
      title: null,
    });

    expect(chart?.title).toBeUndefined();
    expect(store.getState().charts.charts[0]?.title).toBeUndefined();
  });

  it('accepts host placement and visual options', () => {
    const store = createSpreadsheetStore();
    const chart = createSessionChart(store, range(0, 0, 3, 1), {
      id: 'sales',
      kind: 'column',
      title: 'Sales',
      color: '#107c10',
      x: 24,
      y: 40,
      w: 480,
      h: 280,
    });

    expect(chart).toMatchObject({
      id: 'sales',
      title: 'Sales',
      color: '#107c10',
      x: 24,
      y: 40,
      w: 480,
      h: 280,
    });
  });

  it('keeps nested chart options detached from callers and the returned command value', () => {
    const store = createSpreadsheetStore();
    const palette = ['#107c10', '#ed7d31'];
    const legend = { visible: true, position: 'bottom' as const };
    const chart = createSessionChart(store, range(0, 0, 3, 1), {
      kind: 'column',
      palette,
      legend,
      xAxis: { title: 'Month' },
      yAxis: { min: 0, max: 100 },
    });
    if (!chart) throw new Error('expected chart');

    palette[0] = '#000000';
    legend.visible = false;
    if (chart.palette) (chart.palette as string[])[0] = '#ffffff';
    expect(store.getState().charts.charts[0]).toMatchObject({
      palette: ['#107c10', '#ed7d31'],
      legend: { visible: true, position: 'bottom' },
      xAxis: { title: 'Month' },
      yAxis: { min: 0, max: 100 },
    });

    const patchPalette = ['#4472c4'];
    const patch = { palette: patchPalette, legend: { visible: false } };
    updateSessionChart(store, chart.id, patch);
    patchPalette[0] = '#000000';
    expect(store.getState().charts.charts[0]?.palette).toEqual(['#4472c4']);
    expect(store.getState().charts.charts[0]?.legend).toEqual({ visible: false });
  });

  it('maps a labeled table to named category series without summing columns', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      data: {
        ...s.data,
        cells: new Map([
          [
            addrKey({ sheet: 0, row: 0, col: 0 }),
            { value: { kind: 'text', value: 'Month' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 0, col: 1 }),
            { value: { kind: 'text', value: 'North' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 0, col: 2 }),
            { value: { kind: 'text', value: 'South' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 1, col: 0 }),
            { value: { kind: 'text', value: 'Jan' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 1, col: 1 }),
            { value: { kind: 'number', value: 10 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 1, col: 2 }),
            { value: { kind: 'number', value: 20 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 2, col: 0 }),
            { value: { kind: 'text', value: 'Feb' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 2, col: 1 }),
            { value: { kind: 'number', value: 30 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 2, col: 2 }),
            { value: { kind: 'number', value: 40 }, formula: null },
          ],
        ]),
      },
    }));

    expect(sessionChartData(store.getState(), range(0, 0, 2, 2), { kind: 'column' })).toEqual({
      domain: 'category',
      categories: ['Jan', 'Feb'],
      series: [
        { id: 'series-0', name: 'North', values: [10, 30] },
        { id: 'series-1', name: 'South', values: [20, 40] },
      ],
    });
  });

  it('forwards explicit header flags from a session chart to the data adapter', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      data: {
        ...s.data,
        cells: new Map([
          [
            addrKey({ sheet: 0, row: 0, col: 0 }),
            { value: { kind: 'text', value: 'Category' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 0, col: 1 }),
            { value: { kind: 'number', value: 2025 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 0, col: 2 }),
            { value: { kind: 'number', value: 2026 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 1, col: 0 }),
            { value: { kind: 'text', value: 'A' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 1, col: 1 }),
            { value: { kind: 'number', value: 1 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 1, col: 2 }),
            { value: { kind: 'number', value: 2 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 2, col: 0 }),
            { value: { kind: 'text', value: 'B' }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 2, col: 1 }),
            { value: { kind: 'number', value: 3 }, formula: null },
          ],
          [
            addrKey({ sheet: 0, row: 2, col: 2 }),
            { value: { kind: 'number', value: 4 }, formula: null },
          ],
        ]),
      },
    }));

    const chart = createSessionChart(store, range(0, 0, 2, 2), {
      kind: 'column',
      seriesDirection: 'columns',
      headerRow: true,
      headerColumn: true,
    });
    if (!chart) throw new Error('expected chart');
    expect(chart).toMatchObject({ headerRow: true, headerColumn: true });
    expect(sessionChartData(store.getState(), chart)).toEqual({
      domain: 'category',
      categories: ['A', 'B'],
      series: [
        { id: 'series-0', name: '2025', values: [1, 3] },
        { id: 'series-1', name: '2026', values: [2, 4] },
      ],
    });
  });

  it('clears charts by id or intersecting source range', () => {
    const store = createSpreadsheetStore();
    createSessionChart(store, range(0, 0, 3, 1), { id: 'a' });
    createSessionChart(store, range(10, 0, 13, 1), { id: 'b' });

    expect(clearSessionChart(store, 'a')).toBe(true);
    expect(store.getState().charts.charts.map((c) => c.id)).toEqual(['b']);

    expect(clearSessionChartsInRange(store, range(11, 0, 12, 1))).toBe(1);
    expect(store.getState().charts.charts).toHaveLength(0);
  });

  it('updates placement and size by id', () => {
    const store = createSpreadsheetStore();
    createSessionChart(store, range(0, 0, 3, 1), { id: 'a' });

    const updated = updateSessionChart(store, 'a', { x: 32, y: 48, w: 420, h: 260 });

    expect(updated).toMatchObject({ id: 'a', x: 32, y: 48, w: 420, h: 260 });
    expect(store.getState().charts.charts[0]).toMatchObject({
      id: 'a',
      x: 32,
      y: 48,
      w: 420,
      h: 260,
    });
  });

  it('lists, finds, and filters session charts for host object panes', () => {
    const store = createSpreadsheetStore();
    const a = createSessionChart(store, range(0, 0, 3, 1), { id: 'a' });
    const b = createSessionChart(store, range(10, 0, 13, 1), { id: 'b' });

    expect(listSessionCharts(store.getState())).toEqual([a, b]);
    expect(sessionChartById(store.getState(), 'b')).toEqual(b);
    expect(sessionChartById(store.getState(), 'missing')).toBeNull();
    expect(sessionChartsForRange(store.getState(), range(1, 0, 2, 0))).toEqual([a]);
    expect(updateSessionChart(store, 'missing', { title: 'No-op' })).toBeNull();
  });

  it('extracts reusable chart series from rows, columns, and table-like ranges', () => {
    const store = createSpreadsheetStore();
    seedNumbers(store, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 1, col: 0, value: 3 },
      { row: 1, col: 1, value: 4 },
    ]);
    const chart = createSessionChart(store, range(0, 0, 1, 1), { id: 'matrix' });

    expect(sessionChartSeries(store.getState(), range(0, 0, 0, 1))).toEqual([
      { label: 'A', value: 1 },
      { label: 'B', value: 2 },
    ]);
    expect(sessionChartSeries(store.getState(), range(0, 0, 1, 0))).toEqual([
      { label: '1', value: 1 },
      { label: '2', value: 3 },
    ]);
    if (!chart) throw new Error('expected chart to be created');
    expect(sessionChartSeries(store.getState(), chart)).toEqual([
      { label: 'A', value: 4 },
      { label: 'B', value: 6 },
    ]);
  });

  it('returns an empty series for huge chart source ranges', () => {
    const store = createSpreadsheetStore();
    seedNumbers(store, [{ row: 0, col: 0, value: 1 }]);

    expect(sessionChartSeries(store.getState(), range(0, 0, 0, 10_000))).toEqual([]);
    expect(sessionChartSeries(store.getState(), range(0, 0, 10_000, 1))).toEqual([]);
  });

  it('rejects create, update, and clear on protected sheets', () => {
    const store = createSpreadsheetStore();
    const chart = createSessionChart(store, range(0, 0, 3, 1), { id: 'a' });
    expect(chart).not.toBeNull();
    setProtectedSheet(store, 0, true);

    expect(createSessionChart(store, range(1, 0, 4, 1), { id: 'b' })).toBeNull();
    expect(updateSessionChart(store, 'a', { x: 32 })).toBeNull();
    expect(clearSessionChart(store, 'a')).toBe(false);
    expect(clearSessionChartsInRange(store, range(0, 0, 3, 1))).toBe(0);
    expect(listSessionCharts(store.getState())).toEqual([chart]);
  });

  it('records session chart create, update, and clear in history', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    const chart = createSessionChart(store, range(0, 0, 3, 1), { id: 'a' }, history);
    expect(chart).not.toBeNull();
    expect(listSessionCharts(store.getState()).map((c) => c.id)).toEqual(['a']);
    expect(history.undo()).toBe(true);
    expect(listSessionCharts(store.getState())).toEqual([]);
    expect(history.redo()).toBe(true);
    expect(listSessionCharts(store.getState()).map((c) => c.id)).toEqual(['a']);

    expect(updateSessionChart(store, 'a', { x: 10 }, history)).toMatchObject({ x: 10 });
    expect(history.undo()).toBe(true);
    expect(sessionChartById(store.getState(), 'a')?.x).toBeUndefined();

    expect(clearSessionChart(store, 'a', history)).toBe(true);
    expect(listSessionCharts(store.getState())).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(listSessionCharts(store.getState()).map((c) => c.id)).toEqual(['a']);
  });
});
