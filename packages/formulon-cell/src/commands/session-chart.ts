import type { Range } from '../engine/types.js';
import {
  mutators,
  type SessionChart,
  type SessionChartKind,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';

export interface CreateSessionChartOptions {
  id?: string;
  kind?: SessionChartKind;
  /** Pass `null` when the renderer should supply a localized fallback title. */
  title?: string | null;
  color?: string;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

export type SessionChartPatch = Partial<Omit<SessionChart, 'id'>>;

export interface SessionChartSeriesPoint {
  label: string;
  value: number;
}

function defaultChartId(range: Range, kind: SessionChartKind): string {
  return `chart-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}-${kind}`;
}

function defaultTitle(kind: SessionChartKind): string {
  return kind === 'line' ? 'Line chart' : 'Column chart';
}

/** Create or replace a session chart overlay for `range`. This is UI-owned
 *  until the engine exposes writable chart definitions; host apps can call
 *  the same command from ribbons, menus, or custom Quick Analysis surfaces. */
export function createSessionChart(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionChartOptions = {},
): SessionChart {
  const kind = options.kind ?? 'column';
  const chart: SessionChart = {
    id: options.id ?? defaultChartId(range, kind),
    kind,
    source: range,
  };
  if (options.title !== null) chart.title = options.title ?? defaultTitle(kind);
  if (options.color !== undefined) chart.color = options.color;
  if (options.x !== undefined) chart.x = options.x;
  if (options.y !== undefined) chart.y = options.y;
  if (options.w !== undefined) chart.w = options.w;
  if (options.h !== undefined) chart.h = options.h;
  mutators.upsertChart(store, chart);
  return chart;
}

export function listSessionCharts(state: {
  charts: { charts: readonly SessionChart[] };
}): readonly SessionChart[] {
  return state.charts.charts;
}

export function sessionChartById(
  state: { charts: { charts: readonly SessionChart[] } },
  id: string,
): SessionChart | null {
  return state.charts.charts.find((chart) => chart.id === id) ?? null;
}

export function sessionChartsForRange(
  state: { charts: { charts: readonly SessionChart[] } },
  range: Range,
): readonly SessionChart[] {
  return state.charts.charts.filter((chart) => rangesIntersect(chart.source, range));
}

/** Remove a session chart overlay by id. */
export function clearSessionChart(store: SpreadsheetStore, id: string): void {
  mutators.removeChart(store, id);
}

/** Patch placement or visual metadata for an existing session chart. */
export function updateSessionChart(
  store: SpreadsheetStore,
  id: string,
  patch: SessionChartPatch,
): SessionChart | null {
  if (!sessionChartById(store.getState(), id)) return null;
  mutators.updateChart(store, id, patch);
  return sessionChartById(store.getState(), id);
}

/** Remove every session chart whose source range intersects `range`. */
export function clearSessionChartsInRange(store: SpreadsheetStore, range: Range): void {
  mutators.clearChartsInRange(store, range);
}

export function sessionChartSeries(
  state: State,
  chartOrRange: SessionChart | Range,
): readonly SessionChartSeriesPoint[] {
  const range = 'source' in chartOrRange ? chartOrRange.source : chartOrRange;
  if (range.sheet !== state.data.sheetIndex) return [];
  const out: SessionChartSeriesPoint[] = [];
  if (range.r0 === range.r1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      out.push({
        label: columnLabel(col),
        value: cellNumber(state, range.sheet, range.r0, col) ?? 0,
      });
    }
    return out;
  }
  if (range.c0 === range.c1) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      out.push({
        label: String(row + 1),
        value: cellNumber(state, range.sheet, row, range.c0) ?? 0,
      });
    }
    return out;
  }
  for (let col = range.c0; col <= range.c1; col += 1) {
    let sum = 0;
    for (let row = range.r0; row <= range.r1; row += 1) {
      sum += cellNumber(state, range.sheet, row, col) ?? 0;
    }
    out.push({ label: columnLabel(col), value: sum });
  }
  return out;
}

const rangesIntersect = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);

const columnLabel = (n: number): string => {
  let v = n;
  let out = '';
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const cellNumber = (state: State, sheet: number, row: number, col: number): number | null => {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  return cell?.value.kind === 'number' && Number.isFinite(cell.value.value)
    ? cell.value.value
    : null;
};
