import { ChartTitleKind, DrawingObjectKind } from '@libraz/formulon';
import {
  chartDataFromEngine,
  chartSpecFromEngine,
  unsupportedChartReason,
} from '../charts/engine-adapter.js';
import type { ChartData } from '../charts/types.js';
import { mintDrawingId } from '../store/drawing-order.js';
import { mutators, type SessionChart, type SpreadsheetStore } from '../store/store.js';
import { rangeOfRect } from './address.js';
import { drawingAnchorOf } from './drawing-geometry.js';
import { followEngineDrawingOrder } from './images-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { EngineDrawingObject } from './workbook-handle-annotations.js';
import type { EngineChartData, EngineChartInfo } from './workbook-handle-charts.js';

const EMPTY_DATA: ChartData = { domain: 'category', categories: [], series: [] };

/**
 * The store id of each chart of `sheet` by object id: the one the store holds,
 * one of `restored`, or a new one. An id names the chart for as long as it
 * lives, whatever sheet it moves to.
 */
const chartIds = (
  store: SpreadsheetStore,
  sheet: number,
  objects: readonly EngineDrawingObject[],
  restored?: ReadonlyMap<number, string>,
): Map<number, string> => {
  const state = store.getState();
  const ids = new Map(
    state.charts.charts
      .filter((chart) => chart.sheet === sheet)
      .map((chart) => [chart.nativeObjectId, chart.id] as const),
  );
  const taken = new Set(ids.values());
  for (const { objectId } of objects) {
    const id =
      ids.get(objectId) ??
      restored?.get(objectId) ??
      mintDrawingId(state, `chart-native-${sheet}-${objectId}`, taken);
    ids.set(objectId, id);
    taken.add(id);
  }
  return ids;
};

/** Text the placeholder of an unsupported chart shows. */
const placeholderTitle = (
  object: EngineDrawingObject,
  info: EngineChartInfo | null,
  data: EngineChartData | null,
): string | null => {
  if (info?.title?.kind === ChartTitleKind.Text) return info.title.text;
  if (info?.title?.kind === ChartTitleKind.Ref && data?.titleText) return data.titleText;
  return object.name || null;
};

/** Project one engine chart onto the store's shape. */
export function projectChart(
  wb: WorkbookHandle,
  sheet: number,
  object: EngineDrawingObject,
  id: string,
): SessionChart {
  const info = wb.getChart(sheet, object.objectId);
  const engineData = info ? wb.getChartData(sheet, object.objectId) : null;
  const base = {
    id,
    sheet,
    nativeObjectId: object.objectId,
    nativeAnchor: drawingAnchorOf(object),
    data: engineData ? chartDataFromEngine(engineData) : EMPTY_DATA,
  };
  if (!info) {
    return {
      ...base,
      kind: 'unsupported',
      unsupported: { reason: 'unparseable', title: placeholderTitle(object, null, null) },
    };
  }
  const source = info.source
    ? rangeOfRect(info.source.range.sheet, info.source.range.rect)
    : undefined;
  const extra = {
    ...(source ? { source } : {}),
    ...(info.isPivotChart ? { pivotChart: true } : {}),
  };
  const spec = chartSpecFromEngine(info, engineData);
  if (!spec) {
    return {
      ...base,
      ...extra,
      kind: 'unsupported',
      unsupported: {
        reason: unsupportedChartReason(info) ?? 'unparseable',
        title: placeholderTitle(object, info, engineData),
      },
    };
  }
  return { ...spec, ...base, ...extra };
}

const chartObjects = (wb: WorkbookHandle, sheet: number): EngineDrawingObject[] => {
  const charts = wb
    .listDrawingObjects(sheet)
    .filter((object) => object.kind === DrawingObjectKind.Chart);
  // A chartsheet holds exactly the chart the engine names.
  const own = wb.chartsheetChartId(sheet);
  return own === null ? charts : charts.filter((object) => object.objectId === own);
};

/** Rebuild the chart projections of `sheet` and take the engine's drawing order. */
export function hydrateChartsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  hydrateRestoredCharts(wb, store, sheet, new Map());
}

/** {@link hydrateChartsFromEngine}, naming each chart of `restored`, put back
 *  after the store dropped it, by the id it had. */
export function hydrateRestoredCharts(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  restored: ReadonlyMap<number, string>,
): void {
  if (!wb.capabilities.charts) return;
  const objects = chartObjects(wb, sheet);
  const ids = chartIds(store, sheet, objects, restored);
  const charts = objects.map((object) =>
    projectChart(wb, sheet, object, ids.get(object.objectId) as string),
  );
  mutators.replaceSheetCharts(store, sheet, charts);
  followEngineDrawingOrder(wb, store, sheet);
}

/** Re-read every chart and the active sheet's, after a sheet rename, move or
 *  removal changed what the references point at. Run it once the store's sheet
 *  indexes follow the edit (`applySheetStructureEvent`), so each chart is
 *  found on its new sheet and keeps its id. */
export function reloadChartsFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.charts) return;
  const state = store.getState();
  const sheets = new Set([state.data.sheetIndex, ...state.charts.charts.map((c) => c.sheet)]);
  for (const sheet of sheets) {
    if (sheet < wb.sheetCount) hydrateChartsFromEngine(wb, store, sheet);
  }
}

/**
 * Re-read one chart after the engine changed it. `restoredId` names a chart
 * put back after the store dropped it. Returns the new projection, or null when
 * the chart no longer exists and its projection was removed.
 */
export function syncChartFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  objectId: number,
  restoredId?: string,
): SessionChart | null {
  const objects = chartObjects(wb, sheet);
  const object = objects.find((candidate) => candidate.objectId === objectId);
  const restored = restoredId === undefined ? undefined : new Map([[objectId, restoredId]]);
  const projected = object
    ? projectChart(
        wb,
        sheet,
        object,
        chartIds(store, sheet, [object], restored).get(objectId) as string,
      )
    : null;
  const byObject = new Map(
    store
      .getState()
      .charts.charts.filter((chart) => chart.sheet === sheet)
      .map((chart) => [chart.nativeObjectId, chart]),
  );
  if (projected) byObject.set(objectId, projected);
  else byObject.delete(objectId);
  mutators.replaceSheetCharts(
    store,
    sheet,
    objects.flatMap((candidate) => byObject.get(candidate.objectId) ?? []),
  );
  return projected;
}

/**
 * After a recalc, re-read every projected chart. The recalc names only the cells
 * that were written, not the formulas that changed with them, so no chart can be
 * skipped on its references alone.
 */
export function refreshChartsAfterRecalc(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.charts) return;
  const sheets = new Set(store.getState().charts.charts.map((chart) => chart.sheet));
  for (const sheet of sheets) hydrateChartsFromEngine(wb, store, sheet);
}
