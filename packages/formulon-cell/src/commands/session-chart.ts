import { cloneChartSpec } from '../charts/clone.js';
import { chartDataFromTable } from '../charts/data.js';
import type {
  ChartData,
  ChartEmptyCellMode,
  ChartGrouping,
  ChartLegendSpec,
  ChartSeriesDirection,
  ChartSpec,
  ChartTableCell,
} from '../charts/types.js';
import { colLetter } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { rangeArea, rangesIntersect } from '../store/selection-geometry.js';
import {
  mutators,
  type SessionChart,
  type SessionChartKind,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import type { History } from './history.js';
import { interactionControllerFor } from './interaction-controller.js';
import type { InteractionOrigin, OperationIntent } from './interaction-policy.js';
import { isSheetProtected } from './protection.js';
import { recordChartsChange } from './slice-history.js';

export interface CreateSessionChartOptions extends Omit<ChartSpec, 'kind' | 'title'> {
  id?: string;
  kind?: SessionChartKind;
  /** Pass `null` when the renderer should supply a localized fallback title. */
  title?: string | null;
  color?: string;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  grouping?: ChartGrouping;
  seriesDirection?: ChartSeriesDirection;
  emptyCells?: ChartEmptyCellMode;
  scatterStyle?: ChartSpec['scatterStyle'];
  headerRow?: boolean;
  headerColumn?: boolean;
  gapWidth?: ChartSpec['gapWidth'];
  overlap?: ChartSpec['overlap'];
  firstSliceAngle?: ChartSpec['firstSliceAngle'];
  holeSize?: ChartSpec['holeSize'];
  palette?: readonly string[];
  legend?: ChartLegendSpec;
  xAxis?: ChartSpec['xAxis'];
  yAxis?: ChartSpec['yAxis'];
}

export interface ChartMutationAuthorization {
  origin?: InteractionOrigin;
  commandId?: string;
}

const chartIntent = (
  ranges: readonly Range[],
  commandId: string,
  authorization: ChartMutationAuthorization = {},
): OperationIntent => ({
  operation: 'object',
  origin: authorization.origin ?? 'instanceApi',
  commandId: authorization.commandId ?? commandId,
  effects: ranges.map((range) => ({
    kind: 'cells',
    cells: [{ sheet: range.sheet, row: range.r0, col: range.c0 }],
  })),
});

const canMutateChart = (store: SpreadsheetStore, intent: OperationIntent): boolean => {
  for (const effect of intent.effects) {
    if (
      effect.kind === 'cells' &&
      effect.cells.some((cell) => isSheetProtected(store.getState(), cell.sheet))
    )
      return false;
  }
  return interactionControllerFor(store)?.canExecute(intent).allowed ?? true;
};

export type SessionChartPatch = Partial<Omit<SessionChart, 'id'>>;

export interface SessionChartSeriesPoint {
  label: string;
  value: number;
}

const cloneSessionChart = (chart: SessionChart): SessionChart => ({
  ...cloneChartSpec(chart),
  source: { ...chart.source },
});

const MAX_SESSION_CHART_SERIES_CELLS = 10_000;

function defaultChartId(range: Range, kind: SessionChartKind): string {
  return `chart-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}-${kind}`;
}

export type RibbonChartAction = SessionChartKind | 'recommended';

/** Pick a chart kind that fits the selection's shape. One-row selections
 *  default to lines, one-column to bars, narrow categorical lookups to pies,
 *  and everything else to columns. The same rule of thumb is used by every
 *  host (React/Vue/playground) so the "Recommended" toolbar entry feels
 *  identical regardless of framework. */
export function inferRecommendedChartKind(range: Range): SessionChartKind {
  if (range.r0 === range.r1 && range.c1 - range.c0 >= 2) return 'line';
  if (range.c0 === range.c1 && range.r1 - range.r0 >= 2) return 'bar';
  if (range.c1 - range.c0 === 1 && range.r1 - range.r0 <= 6) return 'pie';
  return 'column';
}

export interface CreateRibbonChartFromSelectionOptions {
  store: SpreadsheetStore;
  range: Range;
  /** Chart kind or `'recommended'` to auto-pick via [[inferRecommendedChartKind]]. */
  action: RibbonChartAction;
  history?: History | null;
  /** Prefix used to derive the chart id, so different hosts can spawn parallel
   *  overlays without colliding. Defaults to `'ribbon-chart'`. */
  idPrefix?: string;
}

/** Spawn a session chart overlay anchored to the current selection using the
 *  same positioning heuristics every ribbon host shares. */
export function createRibbonChartFromSelection({
  store,
  range,
  action,
  history = null,
  idPrefix = 'ribbon-chart',
}: CreateRibbonChartFromSelectionOptions): SessionChart | null {
  const kind = action === 'recommended' ? inferRecommendedChartKind(range) : action;
  const count = store.getState().charts.charts.length;
  return createSessionChart(
    store,
    range,
    {
      id: `${idPrefix}-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}-${kind}-${count}`,
      kind,
      title: null,
      x: 340 + (count % 3) * 24,
      y: 96 + (count % 3) * 24,
      w: 360,
      h: 220,
    },
    history,
  );
}

function defaultTitle(kind: SessionChartKind): string {
  if (kind === 'bar') return 'Bar chart';
  if (kind === 'line') return 'Line chart';
  if (kind === 'area') return 'Area chart';
  if (kind === 'pie') return 'Pie chart';
  if (kind === 'doughnut') return 'Doughnut chart';
  if (kind === 'scatter') return 'Scatter chart';
  return 'Column chart';
}

/** Create or replace a session chart overlay for `range`. This is UI-owned
 *  until the engine exposes writable chart definitions; host apps can call
 *  the same command from ribbons, menus, or custom Quick Analysis surfaces. */
export function createSessionChart(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionChartOptions = {},
  history: History | null = null,
  authorization: ChartMutationAuthorization = {},
): SessionChart | null {
  const kind = options.kind ?? 'column';
  const id = options.id ?? defaultChartId(range, kind);
  const previous = sessionChartById(store.getState(), id);
  const intent = chartIntent(
    previous ? [previous.source, range] : [range],
    'chartInsert',
    authorization,
  );
  if (!canMutateChart(store, intent)) return null;
  const chart: SessionChart = {
    id,
    kind,
    source: { ...range },
    ...(options.grouping !== undefined ? { grouping: options.grouping } : {}),
    ...(options.seriesDirection !== undefined ? { seriesDirection: options.seriesDirection } : {}),
    ...(options.emptyCells !== undefined ? { emptyCells: options.emptyCells } : {}),
    ...(options.scatterStyle !== undefined ? { scatterStyle: options.scatterStyle } : {}),
    ...(options.headerRow !== undefined ? { headerRow: options.headerRow } : {}),
    ...(options.headerColumn !== undefined ? { headerColumn: options.headerColumn } : {}),
    ...(options.gapWidth !== undefined ? { gapWidth: options.gapWidth } : {}),
    ...(options.overlap !== undefined ? { overlap: options.overlap } : {}),
    ...(options.firstSliceAngle !== undefined ? { firstSliceAngle: options.firstSliceAngle } : {}),
    ...(options.holeSize !== undefined ? { holeSize: options.holeSize } : {}),
    ...(options.filter !== undefined ? { filter: options.filter } : {}),
    ...(options.dataLabels !== undefined ? { dataLabels: options.dataLabels } : {}),
    ...(options.seriesStyles !== undefined ? { seriesStyles: options.seriesStyles } : {}),
    ...(options.varyColors !== undefined ? { varyColors: options.varyColors } : {}),
    ...(options.palette !== undefined ? { palette: [...options.palette] } : {}),
    ...(options.legend !== undefined ? { legend: { ...options.legend } } : {}),
    ...(options.xAxis !== undefined ? { xAxis: { ...options.xAxis } } : {}),
    ...(options.yAxis !== undefined ? { yAxis: { ...options.yAxis } } : {}),
    ...(options.title !== null ? { title: options.title ?? defaultTitle(kind) } : {}),
    ...(options.color !== undefined ? { color: options.color } : {}),
    ...(options.x !== undefined ? { x: options.x } : {}),
    ...(options.y !== undefined ? { y: options.y } : {}),
    ...(options.w !== undefined ? { w: options.w } : {}),
    ...(options.h !== undefined ? { h: options.h } : {}),
  };
  recordChartsChange(
    history,
    store,
    () => {
      mutators.upsertChart(store, chart);
    },
    intent,
  );
  return cloneSessionChart(chart);
}

export function listSessionCharts(state: {
  charts: { charts: readonly SessionChart[] };
}): readonly SessionChart[] {
  return state.charts.charts.map(cloneSessionChart);
}

export function sessionChartById(
  state: { charts: { charts: readonly SessionChart[] } },
  id: string,
): SessionChart | null {
  const chart = state.charts.charts.find((candidate) => candidate.id === id);
  return chart ? cloneSessionChart(chart) : null;
}

export function sessionChartsForRange(
  state: { charts: { charts: readonly SessionChart[] } },
  range: Range,
): readonly SessionChart[] {
  return state.charts.charts
    .filter((chart) => rangesIntersect(chart.source, range))
    .map(cloneSessionChart);
}

/** Remove a session chart overlay by id. */
export function clearSessionChart(
  store: SpreadsheetStore,
  id: string,
  history: History | null = null,
  authorization: ChartMutationAuthorization = {},
): boolean {
  const chart = sessionChartById(store.getState(), id);
  if (!chart) return false;
  const intent = chartIntent([chart.source], 'chartDelete', authorization);
  if (!canMutateChart(store, intent)) return false;
  recordChartsChange(
    history,
    store,
    () => {
      mutators.removeChart(store, id);
    },
    intent,
  );
  return true;
}

/** Patch placement or visual metadata for an existing session chart. */
export function updateSessionChart(
  store: SpreadsheetStore,
  id: string,
  patch: SessionChartPatch,
  history: History | null = null,
  authorization: ChartMutationAuthorization = {},
): SessionChart | null {
  const chart = sessionChartById(store.getState(), id);
  if (!chart) return null;
  const intent = chartIntent(
    patch.source ? [chart.source, patch.source] : [chart.source],
    'chartUpdate',
    authorization,
  );
  if (!canMutateChart(store, intent)) return null;
  recordChartsChange(
    history,
    store,
    () => {
      mutators.updateChart(store, id, patch);
    },
    intent,
  );
  return sessionChartById(store.getState(), id);
}

/** Remove every session chart whose source range intersects `range`. */
export function clearSessionChartsInRange(
  store: SpreadsheetStore,
  range: Range,
  history: History | null = null,
  authorization: ChartMutationAuthorization = {},
): number {
  const charts = sessionChartsForRange(store.getState(), range);
  const count = charts.length;
  if (count === 0) return 0;
  const intent = chartIntent(
    charts.map((chart) => chart.source),
    'chartDelete',
    authorization,
  );
  if (!canMutateChart(store, intent)) return 0;
  recordChartsChange(
    history,
    store,
    () => {
      mutators.clearChartsInRange(store, range);
    },
    intent,
  );
  return count;
}

/** Build the reusable chart model from the cells covered by a session chart.
 * The adapter deliberately stops at a bounded range so a malformed or
 * accidental whole-sheet chart cannot block the UI. */
export function sessionChartData(
  state: State,
  chartOrRange: SessionChart | Range,
  options: Partial<Pick<ChartSpec, 'kind' | 'seriesDirection' | 'headerRow' | 'headerColumn'>> = {},
): ChartData {
  const chart = ('kind' in chartOrRange ? chartOrRange : undefined) as SessionChart | undefined;
  const range: Range = chart ? chart.source : (chartOrRange as Range);
  const kind = options.kind ?? chart?.kind ?? 'column';
  const seriesDirection = options.seriesDirection ?? chart?.seriesDirection;
  const headerRow = options.headerRow ?? chart?.headerRow;
  const headerColumn = options.headerColumn ?? chart?.headerColumn;
  if (range.sheet !== state.data.sheetIndex || rangeArea(range) > MAX_SESSION_CHART_SERIES_CELLS) {
    return chartDataFromTable([], { kind, seriesDirection, headerRow, headerColumn });
  }

  const table: ChartTableCell[][] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    const values: ChartTableCell[] = [];
    for (let col = range.c0; col <= range.c1; col += 1) {
      const value = state.data.cells.get(`${range.sheet}:${row}:${col}`)?.value;
      if (!value || value.kind === 'blank' || value.kind === 'error') values.push(null);
      else if (value.kind === 'number')
        values.push(Number.isFinite(value.value) ? value.value : null);
      else if (value.kind === 'text') values.push(value.value);
      else values.push(value.value);
    }
    table.push(values);
  }
  return chartDataFromTable(table, { kind, seriesDirection, headerRow, headerColumn });
}

/**
 * @deprecated Use `sessionChartData` for multi-series chart data. This helper
 * retains the original summed-column behavior for compatibility.
 */
export function sessionChartSeries(
  state: State,
  chartOrRange: SessionChart | Range,
): readonly SessionChartSeriesPoint[] {
  const range = 'source' in chartOrRange ? chartOrRange.source : chartOrRange;
  if (range.sheet !== state.data.sheetIndex) return [];
  if (rangeArea(range) > MAX_SESSION_CHART_SERIES_CELLS) return [];
  const out: SessionChartSeriesPoint[] = [];
  if (range.r0 === range.r1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      out.push({
        label: colLetter(col),
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
    out.push({ label: colLetter(col), value: sum });
  }
  return out;
}

const cellNumber = (state: State, sheet: number, row: number, col: number): number | null => {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  return cell?.value.kind === 'number' && Number.isFinite(cell.value.value)
    ? cell.value.value
    : null;
};
