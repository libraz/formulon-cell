import type {
  BlanksAs,
  Chart,
  ChartAxisPatch,
  ChartBarDirection,
  ChartData,
  ChartGrouping,
  ChartInput,
  ChartLegendPosition,
  ChartPlotPatch,
  ChartPlotType,
  ChartPointPatch,
  ChartPreviewOverridesInput,
  ChartProperties,
  ChartSeriesDirection,
  ChartSeriesInput,
  ChartSeriesPatch,
  ChartTitle,
  Color,
  SheetRect,
} from '@libraz/formulon';
import { ColorKind } from '@libraz/formulon';
import { attemptWrite, orFallback } from './result.js';
import { renderedThemeColor } from './theme-color.js';
import type { EngineCapabilities } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { assertAlive, internals } from './workbook-handle-internals.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };

export type EngineChartInfo = ChartProperties;
export type EngineChartData = ChartData;
type EngineChartInsertOptions = ChartInput;

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleChartsMethods {}
}

function chartOf(handle: unknown, sheet: number, objectId: number): Chart {
  return internals(handle).wb.sheet(sheet).chart(objectId);
}

const isThemeColor = (value: object): value is Color =>
  'kind' in value &&
  'theme' in value &&
  'tint' in value &&
  'indexed' in value &&
  value.kind === ColorKind.Theme;

/** `value` with every theme colour rendered to RGB, as the chart paints it. */
function renderThemeColors<T>(value: T, render: (color: Color) => Color): T {
  if (value === null || typeof value !== 'object' || ArrayBuffer.isView(value)) return value;
  if (Array.isArray(value)) return value.map((entry) => renderThemeColors(entry, render)) as T;
  if (isThemeColor(value)) return render(value) as T;
  return Object.fromEntries(
    Object.entries(value).map(([key, entry]) => [key, renderThemeColors(entry, render)]),
  ) as T;
}

export abstract class WorkbookHandleChartsMethods {
  declare readonly capabilities: EngineCapabilities;
  declare resolveColorArgb: WorkbookHandle['resolveColorArgb'];

  /** Typed view of the chart `objectId`, theme colours rendered to RGB; null on failure. */
  getChart(sheet: number, objectId: number): EngineChartInfo | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    const info = orFallback(null, () => chartOf(this, sheet, objectId).getProperties());
    if (!info) return null;
    return renderThemeColors(info, (color) => renderedThemeColor(this, color));
  }

  /** The chart's series resolved against the current cell values. */
  getChartData(sheet: number, objectId: number): EngineChartData | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    return orFallback(null, () => chartOf(this, sheet, objectId).getData());
  }

  /** Resolve the series as if `overrides` were applied; the workbook is unchanged. */
  previewChartData(
    sheet: number,
    objectId: number,
    overrides: ChartPreviewOverridesInput,
  ): EngineChartData | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    return orFallback(null, () => chartOf(this, sheet, objectId).getPreview({ ...overrides }));
  }

  /** Create a chart and return its object id. */
  insertChart(sheet: number, options: EngineChartInsertOptions): number | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    return orFallback(
      null,
      () =>
        internals(this)
          .wb.sheet(sheet)
          .addChart({ ...options }).objectId,
    );
  }

  setChartType(
    sheet: number,
    objectId: number,
    plot: number,
    type: ChartPlotType,
    options: { barDirection?: ChartBarDirection; grouping?: ChartGrouping } = {},
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() => chartOf(this, sheet, objectId).setType(plot, type, options));
  }

  /** Set the title text; null removes the title, an `Auto` title returns it to Excel's automatic one. */
  setChartTitle(sheet: number, objectId: number, title: string | null | ChartTitle): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() => chartOf(this, sheet, objectId).setTitle(title));
  }

  /** Place the legend; a null position removes it. */
  setChartLegend(
    sheet: number,
    objectId: number,
    position: ChartLegendPosition | null,
    overlay: boolean,
    hiddenSeries?: readonly number[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() => {
      const chart = chartOf(this, sheet, objectId);
      if (position === null) chart.setLegend(null);
      else {
        chart.updateLegend({
          position,
          overlay,
          ...(hiddenSeries ? { hiddenEntries: [...hiddenSeries] } : {}),
        });
      }
    });
  }

  setChartAxis(sheet: number, objectId: number, axisId: number, patch: ChartAxisPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() => chartOf(this, sheet, objectId).updateAxis(axisId, { ...patch }));
  }

  /** Set how blanks plot and whether hidden cells are skipped; null removes the element. */
  setChartDisplay(
    sheet: number,
    objectId: number,
    blanksAs: BlanksAs | null,
    plotVisibleOnly: boolean | null,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() =>
      chartOf(this, sheet, objectId).updateDisplay({ blanksAs, plotVisibleOnly }),
    );
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
    return attemptWrite(() => chartOf(this, sheet, objectId).updateSeries(plot, seriesIdx, patch));
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
    return attemptWrite(() => {
      const chart = chartOf(this, sheet, objectId);
      if (patch === null) chart.setPoint(plot, seriesIdx, pointIdx, null);
      else chart.updatePoint(plot, seriesIdx, pointIdx, patch);
    });
  }

  setChartPlot(sheet: number, objectId: number, plot: number, patch: ChartPlotPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() => chartOf(this, sheet, objectId).updatePlot(plot, { ...patch }));
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
    return attemptWrite(() =>
      chartOf(this, sheet, objectId).setFilter({
        hiddenSeries: [...hiddenSeries],
        hiddenCategories: [...hiddenCategories],
      }),
    );
  }

  /** Append a series and return its `c:idx`. */
  addChartSeries(
    sheet: number,
    objectId: number,
    plot: number,
    input: ChartSeriesInput,
  ): number | null {
    assertAlive(this);
    if (!this.capabilities.charts) return null;
    return orFallback(null, () => chartOf(this, sheet, objectId).addSeries(plot, input));
  }

  removeChartSeries(sheet: number, objectId: number, plot: number, seriesIdx: number): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() => chartOf(this, sheet, objectId).removeSeries(plot, seriesIdx));
  }

  setChartSource(
    sheet: number,
    objectId: number,
    source: SheetRect | string,
    direction: ChartSeriesDirection,
    headerRow?: boolean,
    headerColumn?: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.charts) return false;
    return attemptWrite(() =>
      chartOf(this, sheet, objectId).setSource({
        range: source,
        direction,
        ...(headerRow === undefined ? {} : { headerRow }),
        ...(headerColumn === undefined ? {} : { headerColumn }),
      }),
    );
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
