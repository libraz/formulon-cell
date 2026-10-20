import {
  AnchorEditAs,
  AnchorKind,
  type ChartSeriesPatch,
  DrawingObjectKind,
  ChartSeriesDirection as EngineSeriesDirection,
  type ImagePlacement,
} from '@libraz/formulon';
import {
  type ChartEngineOp,
  chartAxisId,
  chartInsertOptions,
  chartSeriesIdx,
  chartSpecFromEngine,
  chartSpecOps,
  withoutSeriesColor,
} from '../charts/engine-adapter.js';
import type { ChartSeriesStyleSpec, ChartSpec } from '../charts/types.js';
import { syncChartFromEngine } from '../engine/charts-sync.js';
import {
  EMU_PER_PX,
  placementForAnchor,
  placementForPixels,
  viewportZoom,
} from '../engine/drawing-geometry.js';
import type { EngineDrawingAnchor, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { EngineChartInfo } from '../engine/workbook-handle-charts.js';
import { cellRectUnclamped, layoutForView } from '../render/geometry.js';
import { rangesIntersect } from '../store/selection-geometry.js';
import { cloneSessionChart } from '../store/sheet-object-mutators.js';
import type { SessionChart, SessionChartKind, SpreadsheetStore } from '../store/store.js';
import { formatSheetAbsoluteRange } from '../wrappers/toolbar-a1.js';
import type { History } from './history.js';
import { interactionControllerFor } from './interaction-controller.js';
import type { InteractionOrigin, OperationIntent } from './interaction-policy.js';
import { isSheetProtected } from './protection.js';
import type { SessionIllustrationArrangeAction } from './session-illustration.js';

export interface ChartMutationAuthorization {
  origin?: InteractionOrigin;
  commandId?: string;
}

export interface CreateSessionChartOptions extends Omit<ChartSpec, 'kind'> {
  /** The workbook the chart is created in. */
  wb: WorkbookHandle;
  kind?: SessionChartKind;
  /** Placement in the engine's units. Omitted, the chart takes Excel's default
   *  frame centred in the visible grid. */
  anchor?: ImagePlacement;
  authorization?: ChartMutationAuthorization;
}

/** Settings to change; `source` re-splits the series from another range and
 *  `anchor` moves or resizes the chart. */
export type SessionChartPatch = Partial<Omit<ChartSpec, 'kind'>> & {
  kind?: SessionChartKind;
  source?: Range;
  anchor?: ImagePlacement;
};

/** The cells of a series to add; each range may lie on any sheet. */
export interface SessionChartSeriesInput {
  /** The cell holding the series name; omitted, the series has none. */
  name?: Range;
  /** The series values: the Y values of a scatter chart. */
  values: Range;
  /** Category labels of a category chart. */
  categories?: Range;
  /** X values of a scatter chart. */
  x?: Range;
}

/** Excel's default chart frame, 360 x 216 pt. */
const DEFAULT_FRAME_CX = 4_572_000;
const DEFAULT_FRAME_CY = 2_743_200;

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

/** Whether protection and the interaction policy allow `intent`. */
export const canMutateChart = (store: SpreadsheetStore, intent: OperationIntent): boolean => {
  for (const effect of intent.effects) {
    if (
      effect.kind === 'cells' &&
      effect.cells.some((cell) => isSheetProtected(store.getState(), cell.sheet))
    )
      return false;
  }
  return interactionControllerFor(store)?.canExecute(intent).allowed ?? true;
};

/** The cells a chart occupies for protection and policy: its anchor cell and its data. */
const chartCells = (chart: SessionChart): Range[] => {
  const { fromRow, fromCol } = chart.nativeAnchor;
  const host = { sheet: chart.sheet, r0: fromRow, c0: fromCol, r1: fromRow, c1: fromCol };
  return chart.source ? [host, chart.source] : [host];
};

/** Push an undo step whose replay re-checks protection and policy first. */
const pushChartHistory = (
  history: History | null,
  store: SpreadsheetStore,
  intent: OperationIntent,
  undo: () => void,
  redo: () => void,
): void => {
  if (!history || history.isReplaying()) return;
  const guarded =
    (origin: 'undo' | 'redo', run: () => void): (() => void) =>
    () => {
      if (!canMutateChart(store, { ...intent, origin })) {
        throw new Error(`chart ${origin} blocked by protection or policy`);
      }
      run();
    };
  history.push({
    undo: guarded('undo', undo),
    redo: guarded('redo', redo),
    intent: { ...intent, origin: 'redo' },
    inverseIntent: { ...intent, origin: 'undo' },
  });
};

/** The `ChartSpec` part of a supported chart's projection. */
const specOf = (chart: SessionChart): ChartSpec => {
  const {
    id: _id,
    sheet: _sheet,
    nativeObjectId: _objectId,
    nativeAnchor: _anchor,
    source: _source,
    unsupported: _unsupported,
    pivotChart: _pivot,
    data: _data,
    ...spec
  } = chart;
  return spec as ChartSpec;
};

/** Series styles without their own colours, so a palette can take effect. */
const withoutSeriesColors = (
  kind: ChartSpec['kind'],
  styles: readonly ChartSeriesStyleSpec[] | undefined,
): ChartSeriesStyleSpec[] | undefined => styles?.map((style) => withoutSeriesColor(kind, style));

/** `current` with the fields of `patch` replaced. A palette replaces the
 *  explicit series colours it would otherwise lose to. */
const patchedSpec = (current: ChartSpec, patch: Partial<ChartSpec>): ChartSpec => {
  const next: ChartSpec = { ...current, ...patch };
  return patch.palette !== undefined && patch.seriesStyles === undefined
    ? { ...next, seriesStyles: withoutSeriesColors(next.kind, current.seriesStyles) }
    : next;
};

/** `style` over `base`; a series colour in `style` replaces every colour `base` stores it in. */
const mergeSeriesStyle = (
  kind: ChartSpec['kind'],
  colored: ChartSeriesStyleSpec,
  style: ChartSeriesStyleSpec,
): ChartSeriesStyleSpec => {
  const base = style.fill?.color === undefined ? colored : withoutSeriesColor(kind, colored);
  return {
    ...base,
    ...style,
    ...(base.fill || style.fill ? { fill: { ...base.fill, ...style.fill } } : {}),
    ...(base.line || style.line ? { line: { ...base.line, ...style.line } } : {}),
    ...(base.marker || style.marker ? { marker: { ...base.marker, ...style.marker } } : {}),
  };
};

/** A created chart's target: the fields `spec` sets, laid over what Excel's
 *  template wrote for the rest, axis and legend fields included. */
const createdSpec = (template: ChartSpec, spec: ChartSpec): ChartSpec => {
  const baseStyles =
    spec.palette === undefined
      ? template.seriesStyles
      : withoutSeriesColors(spec.kind, template.seriesStyles);
  const styles = (baseStyles ?? []).map((base) => {
    const style = spec.seriesStyles?.find((entry) => entry.seriesId === base.seriesId);
    return style ? mergeSeriesStyle(spec.kind, base, style) : base;
  });
  for (const style of spec.seriesStyles ?? []) {
    if (!styles.some((entry) => entry.seriesId === style.seriesId)) styles.push(style);
  }
  // insertChart already split the series as `spec` asks.
  const { seriesDirection: _direction, headerRow: _row, headerColumn: _column, ...settings } = spec;
  const merged: ChartSpec = { ...template, ...settings, seriesStyles: styles };
  return {
    ...merged,
    ...(template.legend || spec.legend ? { legend: { ...template.legend, ...spec.legend } } : {}),
    ...(template.xAxis || spec.xAxis ? { xAxis: { ...template.xAxis, ...spec.xAxis } } : {}),
    ...(template.yAxis || spec.yAxis ? { yAxis: { ...template.yAxis, ...spec.yAxis } } : {}),
  };
};

const runChartOp = (
  wb: WorkbookHandle,
  sheet: number,
  objectId: number,
  info: EngineChartInfo,
  op: ChartEngineOp,
): boolean => {
  switch (op.op) {
    case 'setChartType':
      return wb.setChartType(sheet, objectId, op.plot, op.kind, op.grouping);
    case 'setChartSource':
      return wb.setChartSource(
        sheet,
        objectId,
        op.source,
        op.direction,
        op.headerRow,
        op.headerColumn,
      );
    case 'setChartPlot':
      return wb.setChartPlot(sheet, objectId, op.plot, op.patch);
    case 'setChartDisplay':
      return wb.setChartDisplay(sheet, objectId, op.blanksAs, op.plotVisibleOnly);
    case 'setChartTitle':
      return wb.setChartTitle(sheet, objectId, op.title);
    case 'setChartLegend':
      return wb.setChartLegend(sheet, objectId, op.position, op.overlay);
    case 'setChartAxis': {
      const axisId = chartAxisId(info, op.axis);
      return axisId !== null && wb.setChartAxis(sheet, objectId, axisId, op.patch);
    }
    case 'setChartSeries':
      return wb.setChartSeries(sheet, objectId, op.plot, op.seriesIdx, op.patch);
    case 'setChartPoint':
      return wb.setChartPoint(sheet, objectId, op.plot, op.seriesIdx, op.pointIdx, op.patch);
    case 'setChartFilter':
      return wb.setChartFilter(sheet, objectId, op.hiddenSeries, op.hiddenCategories);
  }
};

/**
 * Bring the engine chart to `target`. A type or source change rebuilds the
 * series, so the chart is re-read after one and the rest is diffed again. The
 * split Excel resolves for `auto` and omitted header flags is taken as reached.
 */
const applyChartSpec = (
  wb: WorkbookHandle,
  sheet: number,
  objectId: number,
  target: ChartSpec,
): boolean => {
  let goal = target;
  const rebuilt = new Set<ChartEngineOp['op']>();
  for (;;) {
    const info = wb.getChart(sheet, objectId);
    const current = info ? chartSpecFromEngine(info, wb.getChartData(sheet, objectId)) : null;
    if (!info || !current) return false;
    let rebuild: ChartEngineOp['op'] | null = null;
    for (const op of chartSpecOps(info, current, goal)) {
      if (op.op === 'setChartType' || op.op === 'setChartSource') {
        // A second rebuild of the same kind means the engine did not reach the goal.
        if (rebuilt.has(op.op)) return false;
        rebuild = op.op;
      }
      if (!runChartOp(wb, sheet, objectId, info, op)) return false;
      if (rebuild) break;
    }
    if (!rebuild) return true;
    rebuilt.add(rebuild);
    if (rebuild === 'setChartSource') {
      const split = wb.getChart(sheet, objectId)?.source;
      if (!split) return false;
      goal = {
        ...goal,
        seriesDirection: split.direction === EngineSeriesDirection.Rows ? 'rows' : 'columns',
        headerRow: split.headerRow,
        headerColumn: split.headerColumn,
      };
    }
  }
};

const sameRange = (a: Range, b: Range | undefined): boolean =>
  b !== undefined &&
  a.sheet === b.sheet &&
  a.r0 === b.r0 &&
  a.c0 === b.c0 &&
  a.r1 === b.r1 &&
  a.c1 === b.c1;

const sourceRef = (wb: WorkbookHandle, range: Range): string =>
  formatSheetAbsoluteRange(wb.sheetName(range.sheet), range);

/** Excel's default frame centred in the visible grid, as a two-cell anchor. */
const defaultPlacement = (store: SpreadsheetStore): ImagePlacement => {
  const state = store.getState();
  const { viewport } = state;
  const fallback: ImagePlacement = {
    anchorKind: AnchorKind.TwoCell,
    editAs: AnchorEditAs.TwoCell,
    row: viewport.rowStart,
    col: viewport.colStart,
    rowOffEmu: 0,
    colOffEmu: 0,
    widthEmu: DEFAULT_FRAME_CX,
    heightEmu: DEFAULT_FRAME_CY,
  };
  try {
    const layout = layoutForView(state);
    const zoom = viewportZoom(store);
    const first = cellRectUnclamped(layout, viewport, viewport.rowStart, viewport.colStart);
    const last = cellRectUnclamped(
      layout,
      viewport,
      viewport.rowStart + Math.max(0, viewport.rowCount - 1),
      viewport.colStart + Math.max(0, viewport.colCount - 1),
    );
    const w = (DEFAULT_FRAME_CX / EMU_PER_PX) * zoom;
    const h = (DEFAULT_FRAME_CY / EMU_PER_PX) * zoom;
    const origin: EngineDrawingAnchor = {
      anchorKind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: viewport.rowStart,
      fromCol: viewport.colStart,
      fromRowOff: 0,
      fromColOff: 0,
      toRow: 0,
      toCol: 0,
      toRowOff: 0,
      toColOff: 0,
      cx: DEFAULT_FRAME_CX,
      cy: DEFAULT_FRAME_CY,
    };
    return placementForPixels(store, origin, {
      x: Math.max(first.x, (first.x + last.x + last.w - w) / 2),
      y: Math.max(first.y, (first.y + last.y + last.h - h) / 2),
      w,
      h,
    });
  } catch {
    return fallback;
  }
};

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
  wb: WorkbookHandle;
  range: Range;
  /** Chart kind or `'recommended'` to auto-pick via [[inferRecommendedChartKind]]. */
  action: RibbonChartAction;
  history?: History | null;
}

/** Insert a chart of the selection the way every ribbon host does. */
export function createRibbonChartFromSelection({
  store,
  wb,
  range,
  action,
  history = null,
}: CreateRibbonChartFromSelectionOptions): SessionChart | null {
  const kind = action === 'recommended' ? inferRecommendedChartKind(range) : action;
  return createSessionChart(store, range, { wb, kind }, history);
}

/**
 * Create a chart of `range` in the workbook and project it into the store.
 * Settings the options leave out take the values Excel's template writes; an
 * omitted title is Excel's automatic one and an empty title is none. Returns
 * null, creating nothing, when the engine has no chart support, the sheet is
 * protected, the policy refuses, or the engine rejects a step.
 */
export function createSessionChart(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionChartOptions,
  history: History | null = null,
): SessionChart | null {
  const { wb, kind = 'column', anchor, authorization, ...settings } = options;
  if (!wb.capabilities.charts) return null;
  const sheet = range.sheet;
  const placement = { ...defaultPlacement(store), ...anchor };
  const row = placement.row ?? 0;
  const col = placement.col ?? 0;
  const intent = chartIntent(
    [{ sheet, r0: row, c0: col, r1: row, c1: col }, range],
    'chartInsert',
    authorization,
  );
  if (!canMutateChart(store, intent)) return null;
  const spec: ChartSpec = { ...settings, kind };
  const objectId = wb.insertChart(sheet, {
    ...chartInsertOptions(spec),
    source: sourceRef(wb, range),
    // Without a title text the template's automatic title stays until the spec is applied.
    ...(typeof spec.title === 'string' ? { title: spec.title } : {}),
    placement,
  });
  if (objectId === null) return null;
  const info = wb.getChart(sheet, objectId);
  const template = info ? chartSpecFromEngine(info, wb.getChartData(sheet, objectId)) : null;
  const snapshot =
    template && applyChartSpec(wb, sheet, objectId, createdSpec(template, spec))
      ? wb.snapshotDrawing(sheet, objectId)
      : null;
  if (!snapshot) {
    wb.removeDrawing(sheet, objectId);
    return null;
  }
  const created = syncChartFromEngine(wb, store, sheet, objectId);
  if (!created) return null;
  pushChartHistory(
    history,
    store,
    intent,
    () => {
      if (!wb.removeDrawing(sheet, objectId)) throw new Error('chart create undo failed');
      syncChartFromEngine(wb, store, sheet, objectId);
    },
    () => {
      if (wb.restoreDrawing(sheet, snapshot) === null) throw new Error('chart create redo failed');
      syncChartFromEngine(wb, store, sheet, objectId);
    },
  );
  return cloneSessionChart(created);
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

/** Charts whose data range intersects `range`; charts with no single data
 *  range are never included. */
export function sessionChartsForRange(
  state: { charts: { charts: readonly SessionChart[] } },
  range: Range,
): readonly SessionChart[] {
  return state.charts.charts
    .filter((chart) => chart.source !== undefined && rangesIntersect(chart.source, range))
    .map(cloneSessionChart);
}

/**
 * Run `edit` on a chart as one undo step that restores the chart as it was.
 * When `edit` fails the chart is restored at once and null is returned.
 */
const editChart = (
  store: SpreadsheetStore,
  chart: SessionChart,
  history: History | null,
  wb: WorkbookHandle,
  intent: OperationIntent,
  label: string,
  edit: () => boolean,
): SessionChart | null => {
  const { sheet, nativeObjectId } = chart;
  const sync = (): SessionChart | null => syncChartFromEngine(wb, store, sheet, nativeObjectId);
  const before = wb.snapshotDrawing(sheet, nativeObjectId);
  if (!before) return null;
  const after = edit() ? wb.snapshotDrawing(sheet, nativeObjectId) : null;
  if (!after) {
    wb.restoreDrawing(sheet, before);
    sync();
    return null;
  }
  const updated = sync();
  pushChartHistory(
    history,
    store,
    intent,
    () => {
      if (wb.restoreDrawing(sheet, before) === null) throw new Error(`chart ${label} undo failed`);
      sync();
    },
    () => {
      if (wb.restoreDrawing(sheet, after) === null) throw new Error(`chart ${label} redo failed`);
      sync();
    },
  );
  return updated && cloneSessionChart(updated);
};

/** Delete one chart and its projection; Undo restores it under the same id. */
const removeChart = (
  store: SpreadsheetStore,
  chart: SessionChart,
  history: History | null,
  wb: WorkbookHandle,
  intent: OperationIntent,
): boolean => {
  const { sheet, nativeObjectId } = chart;
  const snapshot = wb.snapshotDrawing(sheet, nativeObjectId);
  if (!snapshot || !wb.removeDrawing(sheet, nativeObjectId)) return false;
  syncChartFromEngine(wb, store, sheet, nativeObjectId);
  pushChartHistory(
    history,
    store,
    intent,
    () => {
      if (wb.restoreDrawing(sheet, snapshot) === null) throw new Error('chart delete undo failed');
      syncChartFromEngine(wb, store, sheet, nativeObjectId);
    },
    () => {
      if (!wb.removeDrawing(sheet, nativeObjectId)) throw new Error('chart delete redo failed');
      syncChartFromEngine(wb, store, sheet, nativeObjectId);
    },
  );
  return true;
};

/** Delete a chart from the workbook by id. */
export function clearSessionChart(
  store: SpreadsheetStore,
  id: string,
  history: History | null,
  wb: WorkbookHandle,
  authorization: ChartMutationAuthorization = {},
): boolean {
  const chart = sessionChartById(store.getState(), id);
  if (!chart || !wb.capabilities.charts) return false;
  const intent = chartIntent(chartCells(chart), 'chartDelete', authorization);
  if (!canMutateChart(store, intent)) return false;
  return removeChart(store, chart, history, wb, intent);
}

/**
 * Move a chart in its sheet's stacking order among the sheet's charts, the way
 * the spreadsheet's Bring to Front and Send to Back do. Returns null, changing
 * nothing, when the chart is already there or the change is refused.
 */
export function arrangeSessionChart(
  store: SpreadsheetStore,
  id: string,
  action: SessionIllustrationArrangeAction,
  history: History | null,
  wb: WorkbookHandle,
  authorization: ChartMutationAuthorization = {},
): SessionChart | null {
  const chart = sessionChartById(store.getState(), id);
  if (!chart || !wb.capabilities.charts) return null;
  const intent = chartIntent(chartCells(chart), 'chartArrange', authorization);
  if (!canMutateChart(store, intent)) return null;
  const { sheet, nativeObjectId } = chart;
  // Charts keep the slots they hold among the sheet's drawings, so pictures stay put.
  const slots = wb
    .listDrawingObjects(sheet)
    .flatMap((object, index) => (object.kind === DrawingObjectKind.Chart ? [index] : []));
  const rank = wb
    .listDrawingObjects(sheet)
    .filter((object) => object.kind === DrawingObjectKind.Chart)
    .findIndex((object) => object.objectId === nativeObjectId);
  if (rank < 0) return null;
  const targetRank =
    action === 'bring-front'
      ? slots.length - 1
      : action === 'send-back'
        ? 0
        : action === 'bring-forward'
          ? Math.min(slots.length - 1, rank + 1)
          : Math.max(0, rank - 1);
  const from = slots[rank];
  const to = slots[targetRank];
  if (from === undefined || to === undefined || from === to) return null;
  const sync = (): SessionChart | null => syncChartFromEngine(wb, store, sheet, nativeObjectId);
  if (!wb.setDrawingZOrder(sheet, nativeObjectId, to)) return null;
  const arranged = sync();
  pushChartHistory(
    history,
    store,
    intent,
    () => {
      if (!wb.setDrawingZOrder(sheet, nativeObjectId, from)) {
        throw new Error('chart arrange undo failed');
      }
      sync();
    },
    () => {
      if (!wb.setDrawingZOrder(sheet, nativeObjectId, to)) {
        throw new Error('chart arrange redo failed');
      }
      sync();
    },
  );
  return arranged && cloneSessionChart(arranged);
}

/**
 * Change a chart's settings, data range or placement. A pure move or resize
 * re-anchors the chart; any other change is undone by restoring the chart as
 * it was. When the engine rejects a step the chart is restored and null is
 * returned. A chart drawn as a placeholder can only be moved or resized.
 */
export function updateSessionChart(
  store: SpreadsheetStore,
  id: string,
  patch: SessionChartPatch,
  history: History | null,
  wb: WorkbookHandle,
  authorization: ChartMutationAuthorization = {},
): SessionChart | null {
  const chart = sessionChartById(store.getState(), id);
  if (!chart || !wb.capabilities.charts) return null;
  const { source, anchor, ...settings } = patch;
  const intent = chartIntent(
    source ? [...chartCells(chart), source] : chartCells(chart),
    'chartUpdate',
    authorization,
  );
  if (!canMutateChart(store, intent)) return null;
  const { sheet, nativeObjectId } = chart;
  const placement = anchor
    ? {
        anchorKind: chart.nativeAnchor.anchorKind,
        editAs: chart.nativeAnchor.editAs,
        ...anchor,
      }
    : null;
  const sync = (): SessionChart | null => syncChartFromEngine(wb, store, sheet, nativeObjectId);

  if (Object.keys(settings).length === 0 && !source) {
    if (!placement) return chart;
    const before = placementForAnchor(store, chart.nativeAnchor);
    if (!wb.setDrawingAnchor(sheet, nativeObjectId, placement)) return null;
    const moved = sync();
    pushChartHistory(
      history,
      store,
      intent,
      () => {
        if (!wb.setDrawingAnchor(sheet, nativeObjectId, before)) {
          throw new Error('chart move undo failed');
        }
        sync();
      },
      () => {
        if (!wb.setDrawingAnchor(sheet, nativeObjectId, placement)) {
          throw new Error('chart move redo failed');
        }
        sync();
      },
    );
    return moved && cloneSessionChart(moved);
  }

  if (chart.kind === 'unsupported') return null;
  const target = patchedSpec(specOf(chart), settings as Partial<ChartSpec>);
  const split = chartInsertOptions(target);
  return editChart(
    store,
    chart,
    history,
    wb,
    intent,
    'update',
    () =>
      (!placement || wb.setDrawingAnchor(sheet, nativeObjectId, placement)) &&
      (!source ||
        sameRange(source, chart.source) ||
        wb.setChartSource(
          sheet,
          nativeObjectId,
          sourceRef(wb, source),
          split.direction,
          split.headerRow,
          split.headerColumn,
        )) &&
      applyChartSpec(wb, sheet, nativeObjectId, target),
  );
}

/**
 * Append a series to a chart; undo restores the chart as it was. Returns null,
 * changing nothing, when the chart cannot be drawn, the sheet is protected, the
 * policy refuses, or the engine rejects the references.
 */
export function addSessionChartSeries(
  store: SpreadsheetStore,
  id: string,
  series: SessionChartSeriesInput,
  history: History | null,
  wb: WorkbookHandle,
  authorization: ChartMutationAuthorization = {},
): SessionChart | null {
  const chart = sessionChartById(store.getState(), id);
  if (!chart || !wb.capabilities.charts || chart.kind === 'unsupported') return null;
  const { name, values, categories, x } = series;
  const refs = [name, values, categories, x].filter((range) => range !== undefined);
  const intent = chartIntent([...chartCells(chart), ...refs], 'chartUpdate', authorization);
  if (!canMutateChart(store, intent)) return null;
  const patch: ChartSeriesPatch =
    chart.kind === 'scatter'
      ? { y: sourceRef(wb, values), ...(x ? { x: sourceRef(wb, x) } : {}) }
      : {
          values: sourceRef(wb, values),
          ...(categories ? { categories: sourceRef(wb, categories) } : {}),
        };
  if (name) patch.name = sourceRef(wb, name);
  const { sheet, nativeObjectId } = chart;
  return editChart(
    store,
    chart,
    history,
    wb,
    intent,
    'series add',
    () => wb.addChartSeries(sheet, nativeObjectId, 0, patch) !== null,
  );
}

/**
 * Remove the series `seriesId` from a chart; undo restores the chart as it
 * was. A chart keeps its last series: removing it returns null.
 */
export function removeSessionChartSeries(
  store: SpreadsheetStore,
  id: string,
  seriesId: string,
  history: History | null,
  wb: WorkbookHandle,
  authorization: ChartMutationAuthorization = {},
): SessionChart | null {
  const chart = sessionChartById(store.getState(), id);
  const seriesIdx = chartSeriesIdx(seriesId);
  if (!chart || seriesIdx === null || !wb.capabilities.charts || chart.kind === 'unsupported') {
    return null;
  }
  const { sheet, nativeObjectId } = chart;
  const plot = wb.getChart(sheet, nativeObjectId)?.plots[0];
  if (!plot || plot.series.length < 2 || !plot.series.some((entry) => entry.idx === seriesIdx)) {
    return null;
  }
  const intent = chartIntent(chartCells(chart), 'chartUpdate', authorization);
  if (!canMutateChart(store, intent)) return null;
  return editChart(store, chart, history, wb, intent, 'series remove', () =>
    wb.removeChartSeries(sheet, nativeObjectId, 0, seriesIdx),
  );
}

/** Delete every chart whose data range intersects `range`, as one undo step. */
export function clearSessionChartsInRange(
  store: SpreadsheetStore,
  range: Range,
  history: History | null,
  wb: WorkbookHandle,
  authorization: ChartMutationAuthorization = {},
): number {
  const charts = sessionChartsForRange(store.getState(), range);
  if (charts.length === 0 || !wb.capabilities.charts) return 0;
  const intent = chartIntent(charts.flatMap(chartCells), 'chartDelete', authorization);
  if (!canMutateChart(store, intent)) return 0;
  let count = 0;
  history?.begin();
  try {
    for (const chart of charts) {
      if (removeChart(store, chart, history, wb, intent)) count += 1;
    }
  } finally {
    history?.end();
  }
  return count;
}
