import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SparklineKind, type SpreadsheetStore } from '../store/store.js';
import type { SelectionStats } from './aggregate.js';
import { clearFormat } from './format.js';
import { formatAsTable as applyFormatAsTable } from './format-as-table.js';
import { createSessionChart } from './session-chart.js';

/**
 * Quick Analysis — a context-aware action sheet that appears anchored at the
 * selection's bottom-right corner. The host renders the popover; this module
 * provides the pure logic that decides which actions are relevant for a given
 * selection.
 *
 * Sections mirror common spreadsheet apps:
 *   - 書式設定 (Formatting)  — fill / data-bar / icon-set tied to numbers.
 *   - グラフ (Charts)        — session chart overlays.
 *   - 合計 (Totals)          — sum / average etc. for numeric ranges.
 *   - テーブル (Tables)      — Format-As-Table / Pivot stub.
 *   - スパークライン         — line / column / win-loss for numeric ranges.
 */
export type QuickAnalysisGroup = 'formatting' | 'charts' | 'totals' | 'tables' | 'sparklines';

export interface QuickAnalysisAction {
  id: QuickAnalysisActionId;
  group: QuickAnalysisGroup;
  /** Untranslated key — host resolves through i18n strings. */
  labelKey: string;
  /** True when the action is a no-op for the current selection. */
  disabled?: boolean;
}

export type QuickAnalysisActionId =
  | 'format-data-bar'
  | 'format-color-scale'
  | 'format-icon-set'
  | 'format-greater-than'
  | 'format-top-10'
  | 'format-clear'
  | 'totals-sum-row'
  | 'totals-sum-col'
  | 'totals-average-row'
  | 'totals-count-row'
  | 'tables-as-table'
  | 'tables-pivot'
  | 'sparkline-line'
  | 'sparkline-column'
  | 'sparkline-win-loss'
  | 'charts-column'
  | 'charts-line';

export interface QuickAnalysisInput {
  /** Range under the popover anchor — typically the primary selection. */
  range: Range;
  /** Aggregate stats over the same range. Drives "totals" availability. */
  stats: SelectionStats;
  /** True when the host can open a PivotTable creation flow. */
  pivotTableAvailable?: boolean;
  /** True when the host has a chart overlay/renderer active. */
  chartAvailable?: boolean;
}

export interface QuickAnalysisExecuteInput extends QuickAnalysisInput {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  actionId: QuickAnalysisActionId;
}

export type QuickAnalysisExecuteResult =
  | {
      ok: true;
      kind: 'conditional-format' | 'formula' | 'sparkline' | 'chart' | 'table' | 'clear-format';
      count: number;
    }
  | { ok: false; reason: 'disabled' | 'unsupported' | 'out-of-bounds' };

const MAX_COL = 16383;
const MAX_ROW = 1048575;

const colLetter = (n: number): string => {
  let v = n;
  let out = '';
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const rangeRef = (r: Range): string =>
  `${colLetter(r.c0)}${r.r0 + 1}:${colLetter(r.c1)}${r.r1 + 1}`;

/** True when the range covers more than a single cell. */
function isMulti(range: Range): boolean {
  return range.r1 > range.r0 || range.c1 > range.c0;
}

/** Build the menu of actions for the supplied selection. The host orders
 *  them by group; within a group the order returned here is canonical. */
export function buildQuickAnalysisActions(input: QuickAnalysisInput): QuickAnalysisAction[] {
  const { range, stats } = input;
  const out: QuickAnalysisAction[] = [];
  const multi = isMulti(range);
  const hasNumbers = stats.numericCount > 0;

  // Formatting section — color highlights work even on a single cell, but
  // data bars / icon sets need at least two numeric values to interpolate.
  out.push({
    id: 'format-data-bar',
    group: 'formatting',
    labelKey: 'dataBar',
    disabled: stats.numericCount < 2,
  });
  out.push({
    id: 'format-color-scale',
    group: 'formatting',
    labelKey: 'colorScale',
    disabled: stats.numericCount < 2,
  });
  out.push({
    id: 'format-icon-set',
    group: 'formatting',
    labelKey: 'iconSet',
    disabled: stats.numericCount < 3,
  });
  out.push({
    id: 'format-greater-than',
    group: 'formatting',
    labelKey: 'greaterThan',
    disabled: !hasNumbers,
  });
  out.push({
    id: 'format-top-10',
    group: 'formatting',
    labelKey: 'top10',
    disabled: stats.numericCount < 2,
  });
  out.push({ id: 'format-clear', group: 'formatting', labelKey: 'clearFormat' });

  // Totals section.
  out.push({
    id: 'totals-sum-row',
    group: 'totals',
    labelKey: 'sumRow',
    disabled: !hasNumbers || !multi,
  });
  out.push({
    id: 'totals-sum-col',
    group: 'totals',
    labelKey: 'sumCol',
    disabled: !hasNumbers || !multi,
  });
  out.push({
    id: 'totals-average-row',
    group: 'totals',
    labelKey: 'avgRow',
    disabled: !hasNumbers || !multi,
  });
  out.push({ id: 'totals-count-row', group: 'totals', labelKey: 'countRow', disabled: !multi });

  // Tables section.
  out.push({ id: 'tables-as-table', group: 'tables', labelKey: 'formatAsTable', disabled: !multi });
  out.push({
    id: 'tables-pivot',
    group: 'tables',
    labelKey: 'pivotStub',
    disabled: !multi || input.pivotTableAvailable !== true,
  });

  // Sparkline section — only meaningful for a horizontal series.
  const horizontalRun = range.r0 === range.r1 && range.c1 > range.c0;
  out.push({
    id: 'sparkline-line',
    group: 'sparklines',
    labelKey: 'sparkLine',
    disabled: !horizontalRun,
  });
  out.push({
    id: 'sparkline-column',
    group: 'sparklines',
    labelKey: 'sparkColumn',
    disabled: !horizontalRun,
  });
  out.push({
    id: 'sparkline-win-loss',
    group: 'sparklines',
    labelKey: 'sparkWinLoss',
    disabled: !horizontalRun,
  });

  out.push({
    id: 'charts-column',
    group: 'charts',
    labelKey: 'chartColumn',
    disabled: !hasNumbers || !multi || input.chartAvailable !== true,
  });
  out.push({
    id: 'charts-line',
    group: 'charts',
    labelKey: 'chartLine',
    disabled: !hasNumbers || !multi || input.chartAvailable !== true,
  });

  return out;
}

/** Group the action list for the host's section-headers UI. */
export function groupQuickAnalysisActions(
  actions: readonly QuickAnalysisAction[],
): Record<QuickAnalysisGroup, QuickAnalysisAction[]> {
  const out: Record<QuickAnalysisGroup, QuickAnalysisAction[]> = {
    formatting: [],
    charts: [],
    totals: [],
    tables: [],
    sparklines: [],
  };
  for (const a of actions) {
    out[a.group].push(a);
  }
  return out;
}

export function quickAnalysisActionById(
  actions: readonly QuickAnalysisAction[],
  actionId: QuickAnalysisActionId,
): QuickAnalysisAction | null {
  return actions.find((action) => action.id === actionId) ?? null;
}

export function isQuickAnalysisActionEnabled(
  input: QuickAnalysisInput,
  actionId: QuickAnalysisActionId,
): boolean {
  const action = quickAnalysisActionById(buildQuickAnalysisActions(input), actionId);
  return action != null && action.disabled !== true;
}

export function enabledQuickAnalysisActions(input: QuickAnalysisInput): QuickAnalysisAction[] {
  return buildQuickAnalysisActions(input).filter((action) => action.disabled !== true);
}

function addConditionalFormat(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { actionId, range, stats, store } = input;
  if (actionId === 'format-data-bar') {
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range,
      color: '#5b9bd5',
      showValue: true,
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-color-scale') {
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range,
      stops: ['#f8696b', '#ffeb84', '#63be7b'],
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-icon-set') {
    mutators.addConditionalRule(store, { kind: 'icon-set', range, icons: 'traffic3' });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-greater-than') {
    mutators.addConditionalRule(store, {
      kind: 'cell-value',
      range,
      op: '>',
      a: stats.avg,
      apply: { fill: '#ffc7ce', color: '#9c0006' },
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-top-10') {
    mutators.addConditionalRule(store, {
      kind: 'top-bottom',
      range,
      mode: 'top',
      n: 10,
      apply: { fill: '#ffeb9c', color: '#9c6500' },
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  return { ok: false, reason: 'unsupported' };
}

function writeTotalFormulas(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { actionId, range, wb } = input;
  const fn =
    actionId === 'totals-average-row'
      ? 'AVERAGE'
      : actionId === 'totals-count-row'
        ? 'COUNTA'
        : 'SUM';
  let count = 0;

  if (
    actionId === 'totals-sum-row' ||
    actionId === 'totals-average-row' ||
    actionId === 'totals-count-row'
  ) {
    const row = range.r1 + 1;
    if (row > MAX_ROW) return { ok: false, reason: 'out-of-bounds' };
    for (let col = range.c0; col <= range.c1; col += 1) {
      const formula = `=${fn}(${colLetter(col)}${range.r0 + 1}:${colLetter(col)}${range.r1 + 1})`;
      wb.setFormula({ sheet: range.sheet, row, col }, formula);
      count += 1;
    }
    return { ok: true, kind: 'formula', count };
  }

  if (actionId === 'totals-sum-col') {
    const col = range.c1 + 1;
    if (col > MAX_COL) return { ok: false, reason: 'out-of-bounds' };
    for (let row = range.r0; row <= range.r1; row += 1) {
      const formula = `=SUM(${colLetter(range.c0)}${row + 1}:${colLetter(range.c1)}${row + 1})`;
      wb.setFormula({ sheet: range.sheet, row, col }, formula);
      count += 1;
    }
    return { ok: true, kind: 'formula', count };
  }

  return { ok: false, reason: 'unsupported' };
}

function addSparkline(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { actionId, range, store } = input;
  const kind: SparklineKind =
    actionId === 'sparkline-column'
      ? 'column'
      : actionId === 'sparkline-win-loss'
        ? 'win-loss'
        : 'line';
  const col = range.c1 + 1;
  if (col > MAX_COL) return { ok: false, reason: 'out-of-bounds' };
  mutators.setSparkline(
    store,
    { sheet: range.sheet, row: range.r0, col },
    { kind, source: rangeRef(range), showNegative: kind !== 'line' },
  );
  return { ok: true, kind: 'sparkline', count: 1 };
}

function addChart(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { actionId, range, store } = input;
  const kind = actionId === 'charts-line' ? 'line' : 'column';
  const count = store.getState().charts.charts.length;
  createSessionChart(store, range, {
    id: `qa-chart-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}-${kind}`,
    kind,
    title: null,
    x: 320 + (count % 3) * 24,
    y: 72 + (count % 3) * 24,
    w: 360,
    h: 220,
  });
  return { ok: true, kind: 'chart', count: 1 };
}

function formatAsTable(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { range, store } = input;
  applyFormatAsTable(store, range, {
    id: `qa-table-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}`,
  });
  return { ok: true, kind: 'table', count: 1 };
}

function clearAnalysisFormatting(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { store, range } = input;
  clearFormat(store.getState(), store);
  mutators.clearConditionalRulesInRange(store, range);
  mutators.clearSparklinesInRange(store, range);
  mutators.clearChartsInRange(store, range);
  mutators.clearTableOverlaysInRange(store, range);
  return { ok: true, kind: 'clear-format', count: 1 };
}

/** Execute one Quick Analysis action. */
export function executeQuickAnalysisAction(
  input: QuickAnalysisExecuteInput,
): QuickAnalysisExecuteResult {
  if (!isQuickAnalysisActionEnabled(input, input.actionId))
    return { ok: false, reason: 'disabled' };
  if (input.actionId === 'format-clear') return clearAnalysisFormatting(input);
  if (input.actionId.startsWith('format-')) {
    return addConditionalFormat(input);
  }
  if (input.actionId.startsWith('totals-')) return writeTotalFormulas(input);
  if (input.actionId.startsWith('charts-')) return addChart(input);
  if (input.actionId === 'tables-as-table') return formatAsTable(input);
  if (input.actionId.startsWith('sparkline-')) return addSparkline(input);
  return { ok: false, reason: 'unsupported' };
}
