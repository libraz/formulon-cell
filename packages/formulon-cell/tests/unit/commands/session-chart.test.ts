import { describe, expect, it } from 'vitest';
import {
  clearSessionChart,
  clearSessionChartsInRange,
  createSessionChart,
  listSessionCharts,
  sessionChartById,
  sessionChartSeries,
  sessionChartsForRange,
  updateSessionChart,
} from '../../../src/commands/session-chart.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
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

  it('can leave the title unset for localized render fallbacks', () => {
    const store = createSpreadsheetStore();
    const chart = createSessionChart(store, range(0, 0, 2, 0), {
      kind: 'column',
      title: null,
    });

    expect(chart.title).toBeUndefined();
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

  it('clears charts by id or intersecting source range', () => {
    const store = createSpreadsheetStore();
    createSessionChart(store, range(0, 0, 3, 1), { id: 'a' });
    createSessionChart(store, range(10, 0, 13, 1), { id: 'b' });

    clearSessionChart(store, 'a');
    expect(store.getState().charts.charts.map((c) => c.id)).toEqual(['b']);

    clearSessionChartsInRange(store, range(11, 0, 12, 1));
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
    expect(sessionChartSeries(store.getState(), chart)).toEqual([
      { label: 'A', value: 4 },
      { label: 'B', value: 6 },
    ]);
  });
});
