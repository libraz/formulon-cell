import { parseAddrKey } from '../engine/address.js';
import { parseRangeRef, rangeForRef } from '../engine/range-resolver.js';
import { cssOfColorSpec } from '../engine/theme-color.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isRowHidden } from '../store/row-visibility.js';
import type { Sparkline, SparklineGroupSettings, State } from '../store/store.js';
import {
  DEFAULT_SPARKLINE_COLOR,
  type SparklineColors,
  type SparklinePoint,
  sparklineExtent,
  sparklineValues,
} from './painters/controls.js';
import { paintSparkline } from './painters.js';

type SparklineView = Pick<State, 'data' | 'layout'>;

const numberAt = (
  state: SparklineView,
  wb: WorkbookHandle | null,
  sheet: number,
  row: number,
  col: number,
): number | null => {
  const value = wb
    ? wb.getValue({ sheet, row, col })
    : state.data.cells.get(`${sheet}:${row}:${col}`)?.value;
  return value?.kind === 'number' && Number.isFinite(value.value) ? value.value : null;
};

/** Cells of `range` in row-major order as numbers; null for a cell that holds none. */
const readNumbers = (
  state: SparklineView,
  wb: WorkbookHandle | null,
  range: Range,
): { row: number; col: number; value: number | null }[] => {
  const out: { row: number; col: number; value: number | null }[] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      out.push({ row, col, value: numberAt(state, wb, range.sheet, row, col) });
    }
  }
  return out;
};

const rangeOfSource = (
  wb: WorkbookHandle | null,
  source: string,
  hostSheet: number,
): Range | null => {
  if (wb) return rangeForRef(wb, source, hostSheet);
  const parsed = parseRangeRef(source);
  // Without a workbook a sheet name cannot be resolved.
  return parsed && parsed.sheetName === null ? { sheet: hostSheet, ...parsed } : null;
};

/** The points a sparkline plots: its source cells in order, minus the cells in hidden rows
 *  and columns unless the group shows hidden data, on a date axis or by position. Null when
 *  the source does not resolve. */
export function readSparklinePoints(
  state: SparklineView,
  wb: WorkbookHandle | null,
  member: Sparkline,
  hostSheet: number,
  settings: SparklineGroupSettings,
): SparklinePoint[] | null {
  const range = rangeOfSource(wb, member.source, hostSheet);
  if (!range) return null;
  const cells = readNumbers(state, wb, range);
  const dateRange = settings.dateAxis ? settings.dateAxisRange : null;
  const dates = dateRange
    ? readNumbers(state, wb, {
        sheet: dateRange.sheet,
        r0: dateRange.rect.firstRow,
        c0: dateRange.rect.firstCol,
        r1: dateRange.rect.lastRow,
        c1: dateRange.rect.lastCol,
      })
    : null;
  const byDate =
    dates !== null && dates.length === cells.length && dates.every((d) => d.value !== null);
  const visibleSheet = range.sheet === state.data.sheetIndex;
  const points: SparklinePoint[] = [];
  cells.forEach((cell, i) => {
    if (
      !settings.displayHidden &&
      visibleSheet &&
      (isRowHidden(state.layout, cell.row) || state.layout.hiddenCols.has(cell.col))
    ) {
      return;
    }
    points.push({ x: byDate ? (dates[i]?.value as number) : points.length, value: cell.value });
  });
  return points;
}

/** Resolves the min and max over every member of a group, for group-scaled axes; each group
 *  is read once per paint pass. */
export function createSparklineGroupExtents(
  state: State,
  wb: WorkbookHandle | null,
): (groupId: string) => { min: number; max: number } | null {
  const memo = new Map<string, { min: number; max: number } | null>();
  return (groupId) => {
    if (memo.has(groupId)) return memo.get(groupId) ?? null;
    const settings = state.sparkline.groups.get(groupId);
    const values: number[] = [];
    if (settings) {
      for (const [key, member] of state.sparkline.sparklines) {
        const at = parseAddrKey(key);
        if (member.group !== groupId || !at) continue;
        const points = readSparklinePoints(state, wb, member, at.sheet, settings);
        if (points) values.push(...sparklineValues(points, settings.emptyCellsAs));
      }
    }
    const extent = sparklineExtent(values);
    memo.set(groupId, extent);
    return extent;
  };
}

/** The group's colours as CSS through the workbook theme. */
export function sparklineColors(
  settings: SparklineGroupSettings,
  wb: WorkbookHandle | null,
): SparklineColors {
  return {
    series: cssOfColorSpec(wb, settings.colorSeries) ?? DEFAULT_SPARKLINE_COLOR,
    negative: cssOfColorSpec(wb, settings.colorNegative),
    axis: cssOfColorSpec(wb, settings.colorAxis),
    markers: cssOfColorSpec(wb, settings.colorMarkers),
    first: cssOfColorSpec(wb, settings.colorFirst),
    last: cssOfColorSpec(wb, settings.colorLast),
    high: cssOfColorSpec(wb, settings.colorHigh),
    low: cssOfColorSpec(wb, settings.colorLow),
  };
}

/** Cell-paint helper: resolves the source series and draws the sparkline from its group's
 *  settings. Same-sheet refs read from `state.data.cells` when there is no workbook, so the
 *  painter still works in tests with no engine. */
export function paintCellSparkline(
  ctx: CanvasRenderingContext2D,
  bounds: { x: number; y: number; w: number; h: number },
  spec: Sparkline,
  state: State,
  wb: WorkbookHandle | null,
  groupExtent?: (groupId: string) => { min: number; max: number } | null,
  pxPerPt?: number,
): void {
  const settings = state.sparkline.groups.get(spec.group);
  if (!settings) return;
  const points = readSparklinePoints(state, wb, spec, state.data.sheetIndex, settings);
  if (!points) return;
  paintSparkline(ctx, bounds, {
    settings,
    colors: sparklineColors(settings, wb),
    points,
    groupExtent: groupExtent?.(spec.group) ?? null,
    pxPerPt,
  });
}
