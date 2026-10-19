import { ChartTitleKind, DrawingObjectKind } from '@libraz/formulon';
import {
  chartDataFromEngine,
  chartSpecFromEngine,
  unsupportedChartReason,
} from '../charts/engine-adapter.js';
import type { ChartData } from '../charts/types.js';
import { mutators, type SessionChart, type SpreadsheetStore } from '../store/store.js';
import { drawingAnchorOf } from './drawing-geometry.js';
import { parseRangeRef } from './range-resolver.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { EngineDrawingObject } from './workbook-handle-annotations.js';
import type { EngineChartData, EngineChartInfo } from './workbook-handle-charts.js';

const EMPTY_DATA: ChartData = { domain: 'category', categories: [], series: [] };

export const chartIdFor = (sheet: number, objectId: number): string =>
  `chart-native-${sheet}-${objectId}`;

const sheetIndexByName = (wb: WorkbookHandle, name: string): number => {
  const folded = name.toLowerCase();
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    if (wb.sheetName(sheet).toLowerCase() === folded) return sheet;
  }
  return -1;
};

/** The sheet-qualified `ChartInfo.source.range` as a `Range`. */
const sourceRange = (wb: WorkbookHandle, ref: string, hostSheet: number): Range | undefined => {
  const parsed = parseRangeRef(ref);
  if (!parsed) return undefined;
  const sheet = parsed.sheetName === null ? hostSheet : sheetIndexByName(wb, parsed.sheetName);
  if (sheet < 0) return undefined;
  return { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 };
};

/** Text the placeholder of an unsupported chart shows. */
const placeholderTitle = (
  object: EngineDrawingObject,
  info: EngineChartInfo | null,
  data: EngineChartData | null,
): string | null => {
  if (info?.title.kind === ChartTitleKind.Text) return info.title.text;
  if (info?.title.kind === ChartTitleKind.Ref && data?.titleText) return data.titleText;
  return object.name || null;
};

/** Project one engine chart onto the store's shape. */
export function projectChart(
  wb: WorkbookHandle,
  sheet: number,
  object: EngineDrawingObject,
): SessionChart {
  const id = chartIdFor(sheet, object.objectId);
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
  const source = info.source ? sourceRange(wb, info.source.range, sheet) : undefined;
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

const chartObjects = (wb: WorkbookHandle, sheet: number): EngineDrawingObject[] =>
  wb.listDrawingObjects(sheet).filter((object) => object.kind === DrawingObjectKind.Chart);

/** Rebuild the chart projections of `sheet`, in the engine's stacking order. */
export function hydrateChartsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.charts) return;
  const charts = chartObjects(wb, sheet).map((object) => projectChart(wb, sheet, object));
  mutators.replaceSheetCharts(store, sheet, charts);
}

/** Drop every chart projection and re-read the active sheet's, after a sheet
 *  rename, move or removal changed what the references and indexes point at. */
export function reloadChartsFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.charts) return;
  store.setState((s) => (s.charts.charts.length === 0 ? s : { ...s, charts: { charts: [] } }));
  hydrateChartsFromEngine(wb, store, store.getState().data.sheetIndex);
}

/**
 * Re-read one chart after the engine changed it, keeping the sheet's charts in
 * the engine's stacking order. Returns the new projection, or null when the
 * chart no longer exists and its projection was removed.
 */
export function syncChartFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  objectId: number,
): SessionChart | null {
  const objects = chartObjects(wb, sheet);
  const object = objects.find((candidate) => candidate.objectId === objectId);
  const projected = object ? projectChart(wb, sheet, object) : null;
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
