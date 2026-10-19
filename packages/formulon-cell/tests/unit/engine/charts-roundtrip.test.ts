// @vitest-environment node
/**
 * Charts made in the cell survive a save and reload: the spec read back equals
 * `normalizeChartSpec` of the spec the chart was created from.
 *
 * Factor model (coverwise, pairwise, strength 2, seed 1; 96 rows):
 *   plot         column|bar|line|area x clustered|stacked|percent, pie, doughnut,
 *                scatter x marker|line|lineMarker
 *   direction    auto | columns | rows        headerRow / headerColumn  omit | yes | no
 *   emptyCells   omit | gap | zero | span      title   auto | none | text
 *   legend       hidden | top | right | bottom | left
 *   varyColors   on | off                     colors  palette | explicit
 *   seriesLook   a | b                        points  sparse | all
 *   labels       none | shared | perSeries    filter  none | series | category | both
 *   geometry     low | high                   xAxis   a | b | hidden     yAxis  a | b | log
 * Constraints:
 *   scatter      direction != rows, headerRow = headerColumn = omit
 *   pie|doughnut points = all; column|bar with varyColors on: points = all
 *   (Excel writes a format for every point there, so a sparse list cannot round-trip.)
 *
 * Every field a kind reads back is set, so the comparison covers the whole spec.
 * Set FC_CHART_EXPORT_DIR to also write the first workbook of each kind and the
 * edit case there.
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  type ChartNormalizeContext,
  chartSpecFromEngine,
  normalizeChartSpec,
} from '../../../src/charts/engine-adapter.js';
import { accentsFromTheme } from '../../../src/charts/series-colors.js';
import type {
  ChartAxisSpec,
  ChartDataLabelsSpec,
  ChartKind,
  ChartPointStyleSpec,
  ChartSeriesStyleSpec,
  ChartSpec,
} from '../../../src/charts/types.js';
import { History } from '../../../src/commands/history.js';
import { createSessionChart, updateSessionChart } from '../../../src/commands/session-chart.js';
import { hydrateChartsFromEngine } from '../../../src/engine/charts-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SessionChart } from '../../../src/store/store.js';
import { zipEntries, zipText } from './zip-entries.js';

type Row = readonly [
  plot: string,
  direction: 'auto' | 'columns' | 'rows',
  headerRow: Toggle,
  headerColumn: Toggle,
  emptyCells: 'omit' | 'gap' | 'zero' | 'span',
  title: 'auto' | 'none' | 'text',
  legend: 'hidden' | 'top' | 'right' | 'bottom' | 'left',
  varyColors: 'on' | 'off',
  colors: 'palette' | 'explicit',
  seriesLook: 'a' | 'b',
  points: 'sparse' | 'all',
  labels: 'none' | 'shared' | 'perSeries',
  filter: 'none' | 'series' | 'category' | 'both',
  geometry: 'low' | 'high',
  xAxis: 'a' | 'b' | 'hidden',
  yAxis: 'a' | 'b' | 'log',
];
type Toggle = 'omit' | 'yes' | 'no';

// biome-ignore format: one pairwise row per line
const ROWS: readonly Row[] = [
  ['area-clustered', 'rows', 'yes', 'yes', 'zero', 'text', 'bottom', 'on', 'explicit', 'b', 'all', 'none', 'both', 'high', 'b', 'b'],
  ['line-clustered', 'auto', 'no', 'no', 'omit', 'none', 'top', 'off', 'palette', 'a', 'sparse', 'shared', 'series', 'low', 'b', 'a'],
  ['pie', 'columns', 'omit', 'omit', 'gap', 'auto', 'left', 'on', 'palette', 'a', 'all', 'perSeries', 'category', 'high', 'hidden', 'log'],
  ['column-clustered', 'columns', 'no', 'yes', 'span', 'auto', 'hidden', 'off', 'explicit', 'b', 'sparse', 'none', 'none', 'low', 'hidden', 'log'],
  ['bar-clustered', 'auto', 'yes', 'omit', 'omit', 'none', 'right', 'on', 'palette', 'b', 'all', 'perSeries', 'none', 'low', 'a', 'b'],
  ['area-stacked', 'columns', 'omit', 'no', 'span', 'text', 'top', 'off', 'explicit', 'a', 'all', 'shared', 'category', 'high', 'a', 'b'],
  ['line-clustered', 'rows', 'no', 'omit', 'zero', 'auto', 'right', 'on', 'explicit', 'a', 'sparse', 'perSeries', 'series', 'high', 'a', 'a'],
  ['area-percent', 'rows', 'omit', 'yes', 'gap', 'none', 'left', 'off', 'explicit', 'b', 'sparse', 'none', 'category', 'low', 'b', 'a'],
  ['line-percent', 'columns', 'yes', 'omit', 'span', 'text', 'bottom', 'off', 'palette', 'a', 'all', 'shared', 'both', 'low', 'hidden', 'a'],
  ['line-stacked', 'auto', 'no', 'no', 'omit', 'auto', 'hidden', 'on', 'palette', 'a', 'sparse', 'none', 'both', 'high', 'hidden', 'b'],
  ['bar-clustered', 'columns', 'omit', 'no', 'zero', 'none', 'hidden', 'on', 'explicit', 'b', 'all', 'shared', 'series', 'high', 'b', 'log'],
  ['scatter-lineMarker', 'auto', 'omit', 'omit', 'gap', 'text', 'right', 'off', 'explicit', 'a', 'sparse', 'none', 'series', 'low', 'a', 'log'],
  ['bar-percent', 'rows', 'no', 'yes', 'omit', 'text', 'bottom', 'off', 'palette', 'a', 'all', 'perSeries', 'none', 'high', 'a', 'log'],
  ['bar-stacked', 'auto', 'no', 'yes', 'span', 'none', 'left', 'on', 'explicit', 'a', 'all', 'perSeries', 'series', 'high', 'b', 'b'],
  ['doughnut', 'rows', 'yes', 'omit', 'zero', 'auto', 'top', 'on', 'palette', 'b', 'all', 'perSeries', 'both', 'high', 'hidden', 'log'],
  ['column-stacked', 'columns', 'omit', 'omit', 'omit', 'auto', 'bottom', 'on', 'explicit', 'a', 'all', 'shared', 'none', 'high', 'b', 'a'],
  ['column-percent', 'columns', 'omit', 'yes', 'omit', 'text', 'left', 'off', 'explicit', 'b', 'sparse', 'shared', 'both', 'high', 'a', 'log'],
  ['scatter-marker', 'columns', 'omit', 'omit', 'gap', 'none', 'right', 'on', 'explicit', 'b', 'sparse', 'shared', 'both', 'high', 'hidden', 'b'],
  ['scatter-line', 'auto', 'omit', 'omit', 'zero', 'auto', 'left', 'off', 'explicit', 'a', 'all', 'none', 'category', 'high', 'b', 'a'],
  ['line-percent', 'rows', 'yes', 'no', 'gap', 'auto', 'top', 'off', 'explicit', 'a', 'all', 'perSeries', 'none', 'high', 'b', 'log'],
  ['area-clustered', 'columns', 'omit', 'no', 'zero', 'auto', 'left', 'off', 'palette', 'a', 'sparse', 'shared', 'none', 'low', 'a', 'a'],
  ['bar-stacked', 'columns', 'yes', 'omit', 'omit', 'text', 'hidden', 'off', 'palette', 'b', 'sparse', 'none', 'category', 'high', 'hidden', 'a'],
  ['column-stacked', 'auto', 'no', 'no', 'gap', 'none', 'bottom', 'off', 'palette', 'b', 'sparse', 'perSeries', 'category', 'high', 'hidden', 'b'],
  ['bar-percent', 'columns', 'yes', 'omit', 'span', 'none', 'right', 'on', 'explicit', 'b', 'all', 'none', 'category', 'low', 'b', 'b'],
  ['scatter-lineMarker', 'columns', 'omit', 'omit', 'span', 'none', 'hidden', 'on', 'palette', 'b', 'all', 'shared', 'none', 'high', 'b', 'a'],
  ['line-stacked', 'rows', 'yes', 'omit', 'zero', 'none', 'left', 'off', 'explicit', 'b', 'all', 'shared', 'series', 'low', 'a', 'log'],
  ['column-clustered', 'rows', 'yes', 'no', 'omit', 'text', 'hidden', 'on', 'palette', 'a', 'all', 'perSeries', 'category', 'high', 'a', 'a'],
  ['pie', 'auto', 'omit', 'yes', 'zero', 'none', 'top', 'off', 'explicit', 'a', 'all', 'none', 'both', 'low', 'b', 'b'],
  ['doughnut', 'auto', 'omit', 'yes', 'omit', 'none', 'left', 'off', 'explicit', 'a', 'all', 'shared', 'category', 'low', 'a', 'a'],
  ['area-percent', 'columns', 'no', 'omit', 'zero', 'text', 'right', 'on', 'palette', 'a', 'all', 'none', 'category', 'high', 'hidden', 'b'],
  ['area-stacked', 'auto', 'omit', 'omit', 'gap', 'none', 'hidden', 'on', 'palette', 'b', 'sparse', 'shared', 'series', 'low', 'a', 'a'],
  ['column-percent', 'auto', 'yes', 'yes', 'span', 'auto', 'right', 'on', 'explicit', 'a', 'all', 'none', 'both', 'low', 'a', 'a'],
  ['line-clustered', 'columns', 'omit', 'no', 'gap', 'auto', 'bottom', 'off', 'palette', 'b', 'all', 'perSeries', 'series', 'low', 'hidden', 'b'],
  ['column-percent', 'rows', 'no', 'no', 'zero', 'none', 'top', 'off', 'palette', 'a', 'all', 'perSeries', 'series', 'high', 'hidden', 'b'],
  ['line-percent', 'auto', 'omit', 'yes', 'zero', 'none', 'right', 'on', 'palette', 'b', 'sparse', 'none', 'category', 'high', 'a', 'b'],
  ['area-percent', 'auto', 'yes', 'no', 'omit', 'text', 'right', 'on', 'palette', 'a', 'sparse', 'none', 'none', 'low', 'a', 'a'],
  ['area-percent', 'rows', 'yes', 'omit', 'span', 'auto', 'bottom', 'on', 'explicit', 'b', 'sparse', 'perSeries', 'series', 'high', 'hidden', 'log'],
  ['pie', 'rows', 'no', 'no', 'gap', 'text', 'left', 'on', 'palette', 'b', 'all', 'shared', 'none', 'low', 'a', 'a'],
  ['area-stacked', 'rows', 'yes', 'no', 'omit', 'auto', 'left', 'on', 'explicit', 'b', 'sparse', 'perSeries', 'both', 'low', 'b', 'log'],
  ['area-clustered', 'auto', 'no', 'omit', 'span', 'none', 'top', 'on', 'palette', 'b', 'all', 'perSeries', 'category', 'high', 'hidden', 'log'],
  ['column-stacked', 'rows', 'yes', 'no', 'zero', 'auto', 'hidden', 'on', 'explicit', 'a', 'all', 'none', 'series', 'low', 'a', 'b'],
  ['area-stacked', 'auto', 'no', 'no', 'omit', 'text', 'bottom', 'on', 'palette', 'b', 'sparse', 'none', 'both', 'high', 'hidden', 'log'],
  ['scatter-line', 'columns', 'omit', 'omit', 'omit', 'none', 'bottom', 'on', 'explicit', 'b', 'sparse', 'shared', 'series', 'low', 'hidden', 'log'],
  ['column-clustered', 'auto', 'omit', 'omit', 'zero', 'none', 'bottom', 'off', 'palette', 'a', 'sparse', 'shared', 'series', 'low', 'b', 'b'],
  ['bar-clustered', 'rows', 'no', 'yes', 'gap', 'auto', 'top', 'off', 'explicit', 'a', 'sparse', 'none', 'series', 'low', 'hidden', 'b'],
  ['bar-percent', 'auto', 'omit', 'yes', 'zero', 'text', 'left', 'on', 'explicit', 'b', 'all', 'shared', 'series', 'low', 'a', 'a'],
  ['line-clustered', 'rows', 'omit', 'no', 'span', 'text', 'hidden', 'on', 'palette', 'a', 'sparse', 'none', 'category', 'high', 'a', 'b'],
  ['line-stacked', 'auto', 'omit', 'no', 'span', 'none', 'right', 'off', 'explicit', 'b', 'all', 'perSeries', 'none', 'high', 'hidden', 'a'],
  ['bar-stacked', 'rows', 'omit', 'no', 'zero', 'auto', 'bottom', 'off', 'explicit', 'b', 'sparse', 'shared', 'none', 'low', 'a', 'b'],
  ['line-stacked', 'rows', 'no', 'yes', 'gap', 'none', 'bottom', 'off', 'explicit', 'a', 'all', 'perSeries', 'category', 'high', 'b', 'log'],
  ['bar-percent', 'rows', 'no', 'no', 'gap', 'auto', 'hidden', 'off', 'palette', 'a', 'sparse', 'none', 'both', 'high', 'hidden', 'a'],
  ['scatter-line', 'columns', 'omit', 'omit', 'gap', 'text', 'top', 'off', 'palette', 'a', 'all', 'perSeries', 'none', 'low', 'hidden', 'b'],
  ['scatter-marker', 'auto', 'omit', 'omit', 'zero', 'text', 'right', 'on', 'palette', 'b', 'all', 'perSeries', 'none', 'low', 'a', 'a'],
  ['line-clustered', 'auto', 'yes', 'yes', 'omit', 'text', 'hidden', 'on', 'explicit', 'a', 'sparse', 'shared', 'both', 'high', 'a', 'log'],
  ['area-clustered', 'columns', 'omit', 'no', 'gap', 'none', 'hidden', 'on', 'explicit', 'b', 'sparse', 'shared', 'none', 'low', 'a', 'log'],
  ['doughnut', 'columns', 'no', 'no', 'span', 'text', 'right', 'off', 'palette', 'a', 'all', 'none', 'both', 'low', 'a', 'b'],
  ['column-stacked', 'rows', 'yes', 'yes', 'zero', 'none', 'top', 'on', 'palette', 'b', 'all', 'none', 'series', 'high', 'hidden', 'log'],
  ['scatter-marker', 'columns', 'omit', 'omit', 'span', 'text', 'left', 'on', 'palette', 'a', 'sparse', 'none', 'series', 'low', 'b', 'b'],
  ['bar-stacked', 'columns', 'omit', 'yes', 'gap', 'auto', 'right', 'off', 'palette', 'a', 'all', 'perSeries', 'both', 'high', 'b', 'log'],
  ['line-percent', 'rows', 'no', 'no', 'omit', 'none', 'left', 'off', 'palette', 'b', 'sparse', 'none', 'series', 'low', 'a', 'log'],
  ['area-clustered', 'auto', 'no', 'omit', 'gap', 'none', 'right', 'on', 'explicit', 'a', 'sparse', 'perSeries', 'series', 'high', 'hidden', 'a'],
  ['scatter-lineMarker', 'columns', 'omit', 'omit', 'omit', 'auto', 'left', 'off', 'palette', 'b', 'all', 'none', 'both', 'low', 'hidden', 'b'],
  ['area-percent', 'auto', 'omit', 'yes', 'span', 'auto', 'hidden', 'off', 'palette', 'b', 'all', 'shared', 'both', 'high', 'hidden', 'a'],
  ['area-stacked', 'rows', 'omit', 'yes', 'zero', 'none', 'right', 'on', 'explicit', 'a', 'all', 'none', 'none', 'low', 'a', 'a'],
  ['area-clustered', 'auto', 'yes', 'yes', 'omit', 'text', 'right', 'on', 'palette', 'a', 'sparse', 'none', 'series', 'high', 'hidden', 'b'],
  ['pie', 'columns', 'yes', 'omit', 'omit', 'auto', 'hidden', 'on', 'palette', 'a', 'all', 'shared', 'series', 'high', 'b', 'a'],
  ['column-percent', 'auto', 'omit', 'no', 'gap', 'auto', 'hidden', 'off', 'palette', 'b', 'sparse', 'none', 'category', 'high', 'b', 'b'],
  ['area-percent', 'columns', 'omit', 'yes', 'omit', 'auto', 'top', 'on', 'explicit', 'b', 'sparse', 'perSeries', 'both', 'low', 'a', 'a'],
  ['bar-clustered', 'columns', 'no', 'omit', 'zero', 'auto', 'left', 'on', 'palette', 'a', 'all', 'shared', 'category', 'low', 'hidden', 'log'],
  ['scatter-lineMarker', 'auto', 'omit', 'omit', 'zero', 'text', 'top', 'off', 'explicit', 'a', 'all', 'perSeries', 'category', 'high', 'a', 'log'],
  ['scatter-marker', 'auto', 'omit', 'omit', 'omit', 'auto', 'top', 'on', 'palette', 'a', 'all', 'shared', 'category', 'low', 'hidden', 'log'],
  ['doughnut', 'rows', 'omit', 'no', 'omit', 'none', 'hidden', 'off', 'palette', 'b', 'all', 'shared', 'both', 'high', 'b', 'b'],
  ['scatter-line', 'auto', 'omit', 'omit', 'span', 'auto', 'right', 'on', 'palette', 'a', 'all', 'shared', 'both', 'high', 'a', 'a'],
  ['column-percent', 'columns', 'yes', 'omit', 'span', 'text', 'bottom', 'on', 'palette', 'a', 'all', 'shared', 'none', 'low', 'a', 'a'],
  ['line-clustered', 'columns', 'omit', 'yes', 'gap', 'auto', 'left', 'on', 'explicit', 'a', 'all', 'none', 'none', 'high', 'b', 'b'],
  ['column-stacked', 'auto', 'yes', 'omit', 'span', 'text', 'left', 'on', 'palette', 'a', 'all', 'perSeries', 'both', 'high', 'a', 'log'],
  ['bar-clustered', 'columns', 'omit', 'yes', 'span', 'text', 'bottom', 'on', 'palette', 'b', 'all', 'shared', 'both', 'low', 'b', 'a'],
  ['pie', 'auto', 'no', 'no', 'span', 'auto', 'bottom', 'on', 'palette', 'b', 'all', 'none', 'category', 'low', 'a', 'log'],
  ['line-stacked', 'rows', 'omit', 'no', 'zero', 'text', 'top', 'on', 'explicit', 'a', 'all', 'none', 'series', 'low', 'hidden', 'b'],
  ['column-clustered', 'rows', 'no', 'yes', 'omit', 'auto', 'right', 'on', 'palette', 'a', 'all', 'shared', 'both', 'low', 'b', 'b'],
  ['line-percent', 'columns', 'no', 'yes', 'zero', 'none', 'hidden', 'on', 'palette', 'b', 'all', 'perSeries', 'none', 'high', 'a', 'b'],
  ['column-clustered', 'rows', 'omit', 'yes', 'gap', 'auto', 'top', 'on', 'palette', 'a', 'all', 'perSeries', 'category', 'high', 'hidden', 'b'],
  ['bar-percent', 'columns', 'yes', 'no', 'zero', 'auto', 'top', 'off', 'palette', 'a', 'sparse', 'shared', 'none', 'high', 'a', 'log'],
  ['bar-stacked', 'auto', 'omit', 'yes', 'gap', 'auto', 'top', 'off', 'palette', 'b', 'sparse', 'perSeries', 'category', 'high', 'hidden', 'a'],
  ['line-stacked', 'columns', 'omit', 'yes', 'gap', 'text', 'left', 'off', 'explicit', 'a', 'sparse', 'perSeries', 'category', 'high', 'b', 'a'],
  ['scatter-lineMarker', 'auto', 'omit', 'omit', 'omit', 'text', 'bottom', 'on', 'explicit', 'b', 'sparse', 'none', 'both', 'high', 'b', 'log'],
  ['pie', 'columns', 'yes', 'yes', 'omit', 'auto', 'right', 'off', 'palette', 'b', 'all', 'none', 'series', 'low', 'hidden', 'a'],
  ['scatter-marker', 'auto', 'omit', 'omit', 'zero', 'none', 'hidden', 'on', 'explicit', 'b', 'sparse', 'none', 'series', 'high', 'b', 'b'],
  ['doughnut', 'columns', 'no', 'yes', 'omit', 'none', 'top', 'off', 'palette', 'b', 'all', 'none', 'none', 'low', 'hidden', 'a'],
  ['column-stacked', 'columns', 'omit', 'yes', 'span', 'text', 'right', 'on', 'palette', 'a', 'all', 'perSeries', 'category', 'low', 'hidden', 'a'],
  ['column-clustered', 'rows', 'omit', 'omit', 'zero', 'auto', 'left', 'on', 'palette', 'a', 'all', 'perSeries', 'series', 'high', 'hidden', 'b'],
  ['doughnut', 'rows', 'no', 'no', 'span', 'none', 'bottom', 'on', 'palette', 'b', 'all', 'perSeries', 'none', 'low', 'b', 'a'],
  ['scatter-marker', 'auto', 'omit', 'omit', 'gap', 'auto', 'bottom', 'off', 'palette', 'b', 'all', 'none', 'both', 'high', 'b', 'a'],
  ['doughnut', 'auto', 'omit', 'yes', 'omit', 'auto', 'top', 'off', 'palette', 'b', 'all', 'perSeries', 'series', 'low', 'hidden', 'b'],
  ['doughnut', 'columns', 'omit', 'yes', 'gap', 'text', 'left', 'on', 'palette', 'a', 'all', 'shared', 'none', 'low', 'b', 'log'],
  ['scatter-line', 'auto', 'omit', 'omit', 'span', 'auto', 'hidden', 'off', 'explicit', 'a', 'all', 'none', 'both', 'high', 'b', 'log'],
];

const EXPORT_DIR = process.env.FC_CHART_EXPORT_DIR;

/** A2:A5 label the months, B1:C1 the regions; B3 is blank. */
const CATEGORY_TABLE: (string | number | null)[][] = [
  [null, 'North', 'South'],
  ['Jan', 10, 20],
  ['Feb', null, 25],
  ['Mar', 12, 22],
  ['Apr', 18, 30],
];
/** E1:G5: an x column and two y columns under a header row. */
const SCATTER_TABLE: (string | number | null)[][] = [
  ['x', 'y1', 'y2'],
  [1, 10, 20],
  [2, null, 25],
  [3, 12, 22],
  [4, 18, 30],
];
const CATEGORY_RANGE: Range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 };
const SCATTER_RANGE: Range = { sheet: 0, r0: 0, c0: 4, r1: 4, c1: 6 };

const COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'];

const writeTable = (wb: WorkbookHandle, table: (string | number | null)[][], col0: number) => {
  table.forEach((cells, row) => {
    cells.forEach((value, offset) => {
      const addr = { sheet: 0, row, col: col0 + offset };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else if (typeof value === 'string') wb.setText(addr, value);
    });
  });
};

/** The workbook theme's accent1..accent6, which series markers default to. */
const accentsOf = (wb: WorkbookHandle): string[] => accentsFromTheme(wb.getTheme()?.colors) ?? [];

const toggle = (value: Toggle): boolean | undefined =>
  value === 'omit' ? undefined : value === 'yes';

/** The split Excel's rules give the tables above, as `normalizeChartSpec` takes it. */
const contextFor = (
  kind: ChartKind,
  direction: Row[1],
  headerRow: boolean | undefined,
  headerColumn: boolean | undefined,
): Omit<ChartNormalizeContext, 'accents' | 'styled'> & { readonly points: number } => {
  if (kind === 'scatter') {
    return {
      direction: 'columns',
      headerRow: true,
      headerColumn: false,
      seriesIds: ['series-0', 'series-1'],
      points: 4,
    };
  }
  // Five rows by three columns read as series in columns; the blank corner marks both headers.
  const resolved = direction === 'auto' ? 'columns' : direction;
  const row = headerRow ?? true;
  const column = headerColumn ?? true;
  const series = resolved === 'columns' ? 3 - (column ? 1 : 0) : 5 - (row ? 1 : 0);
  return {
    direction: resolved,
    headerRow: row,
    headerColumn: column,
    seriesIds: Array.from({ length: series }, (_, index) => `series-${index}`),
    points: resolved === 'columns' ? 5 - (row ? 1 : 0) : 3 - (column ? 1 : 0),
  };
};

const pointStyle = (kind: ChartKind, pointIdx: number, markers: boolean): ChartPointStyleSpec => ({
  pointId: kind === 'scatter' ? `point-${pointIdx}` : `category-${pointIdx}`,
  fill: { color: COLORS[(pointIdx + 2) % COLORS.length] },
  line: { color: '#ffffff', width: 1.5 },
  ...(markers ? { marker: { visible: true, shape: 'square', size: 9 } } : {}),
});

const axisSpec = (kind: ChartKind, role: 'x' | 'y', variant: string): ChartAxisSpec => {
  const valueAxis = role === 'y' || kind === 'scatter';
  const base: ChartAxisSpec = {
    visible: variant !== 'hidden',
    reverseOrder: variant === 'b',
    labelPosition: variant === 'b' ? (role === 'x' ? 'low' : 'high') : 'nextTo',
    crossing: variant === 'b' ? (role === 'x' ? 'max' : 10) : 'auto',
    majorGridlines: role === 'y' ? variant !== 'b' : variant === 'b',
    ...(variant === 'a' ? { title: role === 'x' ? 'Month' : 'Sales' } : {}),
  };
  if (!valueAxis) return base;
  if (variant === 'log') return { ...base, logBase: 10, min: 1, max: 100, title: 'Log scale' };
  if (variant === 'hidden') return base;
  return variant === 'a'
    ? { ...base, min: 0, max: 40, majorUnit: 10, numberFormat: '#,##0' }
    : { ...base, min: 5, max: 35, majorUnit: 5, numberFormat: '0.0' };
};

interface RoundTripCase {
  readonly kind: ChartKind;
  readonly range: Range;
  readonly spec: ChartSpec;
  readonly context: Omit<ChartNormalizeContext, 'accents' | 'styled'>;
}

const caseFor = (row: Row): RoundTripCase => {
  const [plot, direction, headerRowKey, headerColumnKey, emptyCells, title, legend] = row;
  const [, , , , , , , varyColors, colors, look, points, labels, filter, geometry, x, y] = row;
  const [kindName, variant] = plot.split('-') as [string, string | undefined];
  const kind = kindName as ChartKind;
  const headerRow = toggle(headerRowKey);
  const headerColumn = toggle(headerColumnKey);
  const context = contextFor(kind, direction, headerRow, headerColumn);
  const scatterStyle =
    kind === 'scatter' ? (variant as 'marker' | 'line' | 'lineMarker') : undefined;
  const markers = kind === 'line' || kind === 'scatter';
  const markersShown = scatterStyle ? scatterStyle !== 'line' : look === 'a';
  const allPoints = Array.from({ length: context.points }, (_, index) => index);
  const seriesStyles = context.seriesIds.map((seriesId, position): ChartSeriesStyleSpec => {
    const pointStyles =
      points === 'all'
        ? allPoints.map((index) => pointStyle(kind, index, markers))
        : position === 0
          ? [pointStyle(kind, 1, markers)]
          : undefined;
    const ownLabels: ChartDataLabelsSpec | undefined =
      labels !== 'perSeries'
        ? undefined
        : position === 0
          ? { visible: true, showSeriesName: true, showValue: false }
          : { visible: true, showValue: true, separator: '; ', numberFormat: '0.0' };
    return {
      seriesId,
      fill: {
        ...(colors === 'explicit' ? { color: COLORS[position % COLORS.length] } : {}),
        ...(look === 'a' ? { transparency: 20 } : {}),
      },
      // A marker-only scatter series draws no line; its line settings read back as absent.
      line: {
        ...(scatterStyle === 'marker' ? {} : { color: look === 'a' ? '#333333' : '#777777' }),
        ...(look === 'a' ? { width: 1.5, dash: 'dash' } : { width: 0.75, dash: 'solid' }),
      },
      ...(markers
        ? {
            // A hidden marker's size and colours read back as absent.
            marker: {
              visible: markersShown,
              ...(markersShown ? { shape: 'diamond' as const } : {}),
              size: 7,
              fill: { color: '#aa0000' },
              line: { color: '#000000', width: 1 },
            },
          }
        : {}),
      ...(ownLabels ? { dataLabels: ownLabels } : {}),
      ...(pointStyles ? { points: pointStyles } : {}),
    };
  });
  const hiddenCategory = kind === 'scatter' ? 'point-1' : 'category-1';
  const grouping =
    variant === 'stacked' ? 'stacked' : variant === 'percent' ? 'percent-stacked' : 'clustered';
  const spec: ChartSpec = {
    kind,
    ...(scatterStyle ? { scatterStyle } : { grouping }),
    seriesDirection: direction,
    ...(headerRow === undefined ? {} : { headerRow }),
    ...(headerColumn === undefined ? {} : { headerColumn }),
    ...(emptyCells === 'omit' ? {} : { emptyCells }),
    ...(title === 'auto' ? {} : { title: title === 'none' ? '' : 'Quarterly sales' }),
    legend: legend === 'hidden' ? { visible: false } : { visible: true, position: legend },
    varyColors: varyColors === 'on',
    ...(colors === 'palette' ? { palette: COLORS } : {}),
    seriesStyles,
    ...(labels === 'shared'
      ? {
          dataLabels: {
            visible: true,
            showValue: true,
            showCategoryName: true,
            separator: ', ',
            numberFormat: '0.0',
            color: '#404040',
            fontSize: 9,
          },
        }
      : {}),
    ...(filter === 'none'
      ? {}
      : {
          filter: {
            ...(filter !== 'category' ? { hiddenSeriesIds: ['series-1'] } : {}),
            ...(filter !== 'series' ? { hiddenCategoryIds: [hiddenCategory] } : {}),
          },
        }),
    gapWidth: geometry === 'low' ? 40 : 300,
    overlap: geometry === 'low' ? -30 : 60,
    firstSliceAngle: geometry === 'low' ? 45 : 300,
    holeSize: geometry === 'low' ? 25 : 80,
    xAxis: axisSpec(kind, 'x', x),
    yAxis: axisSpec(kind, 'y', y),
  };
  return { kind, range: kind === 'scatter' ? SCATTER_RANGE : CATEGORY_RANGE, spec, context };
};

/** `spec` with its series styles in id order. Styles are looked up by id, and the
 *  engine lists a filtered series after the visible ones. */
const byId = (spec: ChartSpec | null): ChartSpec | null =>
  spec?.seriesStyles
    ? {
        ...spec,
        seriesStyles: [...spec.seriesStyles].sort((a, b) =>
          a.seriesId.localeCompare(b.seriesId, 'en', { numeric: true }),
        ),
      }
    : spec;

/** The single chart of sheet 0 after a save and reload, as a `ChartSpec`. */
const reloaded = async (
  wb: WorkbookHandle,
): Promise<{ bytes: Uint8Array; spec: ChartSpec | null; objectId: number | undefined }> => {
  const bytes = wb.save();
  const loaded = await WorkbookHandle.loadBytes(bytes);
  try {
    const [object] = loaded.listDrawingObjects(0).filter((candidate) => candidate.kind === 2);
    if (!object) return { bytes, spec: null, objectId: undefined };
    const info = loaded.getChart(0, object.objectId);
    const data = loaded.getChartData(0, object.objectId);
    return {
      bytes,
      spec: info ? chartSpecFromEngine(info, data) : null,
      objectId: object.objectId,
    };
  } finally {
    loaded.dispose();
  }
};

const exported = new Set<ChartKind>();
const exportWorkbook = (name: string, bytes: Uint8Array): void => {
  if (!EXPORT_DIR) return;
  mkdirSync(EXPORT_DIR, { recursive: true });
  writeFileSync(join(EXPORT_DIR, name), bytes);
};

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe.skipIf(!canLoadWasm())('chart round trip through save and load', () => {
  it.each(ROWS.map((row, index) => [index, row] as const))('row %i', async (index, row) => {
    const { kind, range, spec, context } = caseFor(row);
    const wb = await WorkbookHandle.createDefault();
    try {
      writeTable(wb, CATEGORY_TABLE, 0);
      writeTable(wb, SCATTER_TABLE, 4);
      const store = createSpreadsheetStore();
      const { kind: _kind, ...settings } = spec;
      const chart = createSessionChart(store, range, { wb, kind, ...settings });
      expect(chart, `row ${index}: the engine rejected the chart`).not.toBeNull();

      const { bytes, spec: read } = await reloaded(wb);
      expect(byId(read)).toEqual(
        byId(normalizeChartSpec(spec, { ...context, accents: accentsOf(wb), styled: false })),
      );
      if (!exported.has(kind)) {
        exported.add(kind);
        exportWorkbook(`chart-${kind}.xlsx`, bytes);
      }
    } finally {
      wb.dispose();
    }
  });
});

/** Every series-level fill of `chartXml`: those outside a line, marker, point and labels. */
const seriesFills = (chartXml: string): string[] =>
  [...chartXml.matchAll(/<c:ser>.*?<\/c:ser>/gs)].flatMap(
    ([ser]) =>
      ser
        .replace(/<c:(marker|dPt|dLbls|dLbl)>.*?<\/c:\1>/gs, '')
        .replace(/<a:ln\b.*?<\/a:ln>/gs, '')
        .match(/<c:spPr>.*?(?:Fill|Fill\/)>.*?<\/c:spPr>/gs) ?? [],
  );

describe.skipIf(!canLoadWasm())('line and scatter series colours', () => {
  // Every field the template writes is set, so nothing of the template's accent colours shows.
  const shown = { visible: true, shape: 'circle', size: 7, line: { width: 1 } } as const;
  const line = { width: 2 } as const;
  const pointColor = { fill: { color: '#00ff00', transparency: 50 } } as const;
  it.each([
    [
      'line',
      CATEGORY_RANGE,
      {
        grouping: 'clustered',
        palette: COLORS,
        seriesStyles: [
          {
            seriesId: 'series-0',
            fill: { transparency: 30 },
            line,
            smooth: true,
            marker: shown,
            points: [{ pointId: 'category-1', ...pointColor }],
          },
          {
            seriesId: 'series-1',
            fill: { color: '#123456' },
            line: { ...line, transparency: 25 },
            marker: { visible: false },
          },
        ],
      },
    ],
    [
      'scatter',
      SCATTER_RANGE,
      {
        scatterStyle: 'lineMarker',
        palette: COLORS,
        seriesStyles: [
          {
            seriesId: 'series-0',
            fill: { color: '#123456', transparency: 30 },
            line,
            smooth: true,
            marker: { ...shown, fill: { transparency: 40 } },
            points: [{ pointId: 'point-2', ...pointColor }],
          },
          { seriesId: 'series-1', line, marker: shown },
        ],
      },
    ],
    [
      'scatter',
      SCATTER_RANGE,
      {
        scatterStyle: 'marker',
        palette: COLORS,
        seriesStyles: [
          { seriesId: 'series-0', marker: shown, points: [{ pointId: 'point-1', ...pointColor }] },
          { seriesId: 'series-1', marker: shown },
        ],
      },
    ],
  ] as const)('keeps a %s series colour in the line and markers', async (kind, range, settings) => {
    const wb = await WorkbookHandle.createDefault();
    try {
      writeTable(wb, CATEGORY_TABLE, 0);
      writeTable(wb, SCATTER_TABLE, 4);
      const store = createSpreadsheetStore();
      const spec = { kind, ...settings } as ChartSpec;
      expect(createSessionChart(store, range, { wb, ...spec })).not.toBeNull();
      const { bytes, spec: read } = await reloaded(wb);
      const context = contextFor(kind, 'auto', undefined, undefined);
      expect(byId(read)?.seriesStyles).toEqual(
        byId(normalizeChartSpec(spec, { ...context, accents: accentsOf(wb), styled: false }))
          ?.seriesStyles,
      );
      const entries = zipEntries(bytes);
      const charts = [...entries.keys()].filter((path) => /^xl\/charts\/chart\d+\.xml$/.test(path));
      expect(charts).toHaveLength(1);
      expect(seriesFills(zipText(entries, charts[0] ?? ''))).toEqual([]);
    } finally {
      wb.dispose();
    }
  });
});

const FIXTURE = join(__dirname, '../../fixtures/charts/chart_column.xlsx');

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

describe.skipIf(!canLoadWasm())('editing an Excel chart in the cell', () => {
  it('saves every changed setting, keeps what it did not change, and survives undo and redo', async () => {
    const original = new Uint8Array(readFileSync(FIXTURE));
    const wb = await WorkbookHandle.loadBytes(original);
    try {
      const store = createSpreadsheetStore();
      const history = new History();
      hydrateChartsFromEngine(wb, store, 0);
      const chart = store
        .getState()
        .charts.charts.find(
          (candidate) => candidate.source?.c1 === 2 && candidate.kind === 'column',
        );
      if (!chart) throw new Error('fixture chart missing');

      const rows = updateSessionChart(store, chart.id, { seriesDirection: 'rows' }, history, wb);
      if (!rows) throw new Error('direction change rejected');
      const target: ChartSpec = {
        ...specOf(rows),
        grouping: 'stacked',
        title: 'Edited',
        legend: { visible: true, position: 'left' },
        gapWidth: 90,
        overlap: 100,
        varyColors: false,
        emptyCells: 'zero',
        dataLabels: { visible: true, showValue: true, numberFormat: '0' },
        filter: { hiddenSeriesIds: ['series-1'], hiddenCategoryIds: ['category-0'] },
        palette: ['#123456', '#654321', '#abcdef', '#fedcba'],
        xAxis: { ...rows.xAxis, title: 'Region', reverseOrder: true },
        yAxis: {
          ...rows.yAxis,
          title: { auto: true },
          min: 0,
          max: 60,
          majorUnit: 15,
          numberFormat: '0',
          majorGridlines: false,
        },
      };
      const { kind: _kind, ...patch } = target;
      expect(updateSessionChart(store, chart.id, patch, history, wb)).not.toBeNull();
      expect(history.undo()).toBe(true);
      expect(history.undo()).toBe(true);
      expect(history.redo()).toBe(true);
      expect(history.redo()).toBe(true);

      const { bytes, spec, objectId } = await reloaded(wb);
      exportWorkbook('chart-edit.xlsx', bytes);
      expect(objectId).toBe(chart.nativeObjectId);
      expect(byId(spec)).toEqual(
        byId(
          normalizeChartSpec(target, {
            direction: 'rows',
            headerRow: true,
            headerColumn: true,
            seriesIds: ['series-0', 'series-1', 'series-2', 'series-3'],
            accents: accentsOf(wb),
            // The fixture's charts carry Excel's chart style part.
            styled: true,
          }),
        ),
      );

      // The other charts, the edited chart's style parts and its unmodelled
      // chart-level elements are what Excel wrote.
      const before = zipEntries(original);
      const after = zipEntries(bytes);
      const parts = [...before.keys()].filter((path) => path.startsWith('xl/charts/'));
      expect([...after.keys()].filter((path) => path.startsWith('xl/charts/')).sort()).toEqual(
        parts.sort(),
      );
      const changed = parts.filter(
        (path) =>
          Buffer.compare(
            before.get(path) ?? new Uint8Array(),
            after.get(path) ?? new Uint8Array(),
          ) !== 0,
      );
      expect(changed).toHaveLength(1);
      const edited = zipText(after, changed[0] ?? '');
      expect(edited).toContain('<mc:AlternateContent');
      expect(edited).toContain('c16r3:dataDisplayOptions16');
    } finally {
      wb.dispose();
    }
  });
});
