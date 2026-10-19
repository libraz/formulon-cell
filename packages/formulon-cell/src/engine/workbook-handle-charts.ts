import type {
  ChartAxisPatch,
  ChartBlanksAs,
  ChartData,
  ChartGrouping,
  ChartInfo,
  ChartInsertKind,
  ChartInsertOptions,
  ChartLegendPosition,
  ChartPlotPatch,
  ChartPointPatch,
  ChartPreviewOverrides,
  ChartSeriesDirection,
  ChartSeriesPatch,
  ChartTitleKind,
} from '@libraz/formulon';
import type { EngineCapabilities, Workbook } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

export type EngineChartInfo = ChartInfo;
export type EngineChartData = ChartData;
export type EngineChartInsertOptions = ChartInsertOptions;

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleChartsMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

export abstract class WorkbookHandleChartsMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Typed view of the chart `objectId`; null on failure or capability off. */
  getChart(sheet: number, objectId: number): EngineChartInfo | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    const result = wb(this).getChart(sheet, objectId);
    return result.status.ok ? result.chart : null;
  }

  /** The chart's series resolved against the current cell values. */
  getChartData(sheet: number, objectId: number): EngineChartData | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    const result = wb(this).getChartData(sheet, objectId);
    return result.status.ok ? result.data : null;
  }

  /** Resolve the series as if `overrides` were applied; the workbook is unchanged. */
  previewChartData(
    sheet: number,
    objectId: number,
    overrides: ChartPreviewOverrides,
  ): EngineChartData | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    const result = wb(this).previewChartData(sheet, objectId, { ...overrides });
    return result.status.ok ? result.data : null;
  }

  /** Create a chart and return its object id. */
  insertChart(sheet: number, options: EngineChartInsertOptions): number | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    const result = wb(this).insertChart(sheet, { ...options });
    return result.status.ok ? result.objectId : null;
  }

  setChartType(
    sheet: number,
    objectId: number,
    plot: number,
    kind: ChartInsertKind,
    grouping?: ChartGrouping,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    // The binding rejects an explicit undefined grouping; kinds without one take the 4-argument form.
    const status =
      grouping === undefined
        ? wb(this).setChartType(sheet, objectId, plot, kind)
        : wb(this).setChartType(sheet, objectId, plot, kind, grouping);
    return status.ok;
  }

  /** Set the title text; null removes the title, `{ kind: Auto }` returns it to Excel's automatic one. */
  setChartTitle(
    sheet: number,
    objectId: number,
    title: string | null | { readonly kind: ChartTitleKind },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartTitle(sheet, objectId, title).ok;
  }

  setChartLegend(
    sheet: number,
    objectId: number,
    position: ChartLegendPosition,
    overlay: boolean,
    hiddenSeries?: readonly number[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartLegend(
      sheet,
      objectId,
      position,
      overlay,
      hiddenSeries ? [...hiddenSeries] : undefined,
    ).ok;
  }

  setChartAxis(sheet: number, objectId: number, axisId: number, patch: ChartAxisPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartAxis(sheet, objectId, axisId, { ...patch }).ok;
  }

  setChartDisplay(
    sheet: number,
    objectId: number,
    blanksAs: ChartBlanksAs,
    plotVisibleOnly: boolean | null,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartDisplay(sheet, objectId, blanksAs, plotVisibleOnly).ok;
  }

  setChartSeries(
    sheet: number,
    objectId: number,
    plot: number,
    seriesIdx: number,
    patch: ChartSeriesPatch,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartSeries(sheet, objectId, plot, seriesIdx, patch).ok;
  }

  /** Patch a data point; null removes its formatting. */
  setChartPoint(
    sheet: number,
    objectId: number,
    plot: number,
    seriesIdx: number,
    pointIdx: number,
    patch: ChartPointPatch | null,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartPoint(sheet, objectId, plot, seriesIdx, pointIdx, patch).ok;
  }

  setChartPlot(sheet: number, objectId: number, plot: number, patch: ChartPlotPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartPlot(sheet, objectId, plot, { ...patch }).ok;
  }

  /** Replace the chart filter: series by `c:idx`, categories by position. */
  setChartFilter(
    sheet: number,
    objectId: number,
    hiddenSeries: readonly number[],
    hiddenCategories: readonly number[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartFilter(sheet, objectId, [...hiddenSeries], [...hiddenCategories]).ok;
  }

  /** Append a series and return its `c:idx`. */
  addChartSeries(
    sheet: number,
    objectId: number,
    plot: number,
    patch: ChartSeriesPatch,
  ): number | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    const result = wb(this).addChartSeries(sheet, objectId, plot, patch);
    return result.status.ok ? result.seriesIdx : null;
  }

  removeChartSeries(sheet: number, objectId: number, plot: number, seriesIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).removeChartSeries(sheet, objectId, plot, seriesIdx).ok;
  }

  setChartSource(
    sheet: number,
    objectId: number,
    source: string,
    direction: ChartSeriesDirection,
    headerRow?: boolean,
    headerColumn?: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return wb(this).setChartSource(sheet, objectId, source, direction, headerRow, headerColumn).ok;
  }
}

export function installChartsMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleChartsMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleChartsMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
