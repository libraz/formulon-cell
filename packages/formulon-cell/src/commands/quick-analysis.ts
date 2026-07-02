import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SparklineKind, type SpreadsheetStore } from '../store/store.js';
import type { SelectionStats } from './aggregate.js';
import { formatAsTable as applyFormatAsTable } from './format-as-table.js';
import {
  type History,
  recordChartsChange,
  recordConditionalRulesChange,
  recordFormatChange,
  recordSparklineChange,
  recordTablesChange,
} from './history.js';
import { isCellWritable, isSheetProtected, warnProtected } from './protection.js';
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
 *   - テーブル (Tables)      — Format-As-Table / PivotTable creation flow.
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
  /** Untranslated disabled-reason key for hosts to project into aria/title. */
  disabledReason?: QuickAnalysisDisabledReasonKey;
}

export type QuickAnalysisDisabledReasonKey =
  | 'requiresNumbers'
  | 'requiresTwoNumbers'
  | 'requiresThreeNumbers'
  | 'requiresMultiCell'
  | 'requiresHorizontalRun'
  | 'pivotUnavailable'
  | 'chartUnavailable';

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
  /** Shared undo/redo journal. When supplied, multi-cell Quick Analysis
   *  actions are committed as one Excel-style undo step. */
  history?: History | null;
}

export type QuickAnalysisExecuteResult =
  | {
      ok: true;
      kind: 'conditional-format' | 'formula' | 'sparkline' | 'chart' | 'table' | 'clear-format';
      count: number;
    }
  | { ok: false; reason: 'disabled' | 'unsupported' | 'out-of-bounds' | 'protected' };

const MAX_COL = 16383;
const MAX_ROW = 1048575;
const MAX_QUICK_ANALYSIS_FORMULA_WRITES = 100_000;
const MAX_EXACT_PROTECTION_SCAN_CELLS = 100_000;

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

const rangeArea = (range: Range): number => (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);

const rangeContainsRange = (outer: Range, inner: Range): boolean =>
  outer.sheet === inner.sheet &&
  outer.r0 <= inner.r0 &&
  outer.r1 >= inner.r1 &&
  outer.c0 <= inner.c0 &&
  outer.c1 >= inner.c1;

const addrFromKey = (key: string): { sheet: number; row: number; col: number } | null => {
  const parts = key.split(':').map(Number);
  if (parts.length !== 3) return null;
  const [sheet, row, col] = parts as [number, number, number];
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
};

const rangeContainsAddr = (
  range: Range,
  addr: { sheet: number; row: number; col: number },
): boolean =>
  addr.sheet === range.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

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
  const totalsReason: QuickAnalysisDisabledReasonKey | undefined = !multi
    ? 'requiresMultiCell'
    : !hasNumbers
      ? 'requiresNumbers'
      : undefined;
  const chartReason: QuickAnalysisDisabledReasonKey | undefined = !multi
    ? 'requiresMultiCell'
    : !hasNumbers
      ? 'requiresNumbers'
      : input.chartAvailable !== true
        ? 'chartUnavailable'
        : undefined;

  // Formatting section — color highlights work even on a single cell, but
  // data bars / icon sets need at least two numeric values to interpolate.
  out.push({
    id: 'format-data-bar',
    group: 'formatting',
    labelKey: 'dataBar',
    disabled: stats.numericCount < 2,
    disabledReason: stats.numericCount < 2 ? 'requiresTwoNumbers' : undefined,
  });
  out.push({
    id: 'format-color-scale',
    group: 'formatting',
    labelKey: 'colorScale',
    disabled: stats.numericCount < 2,
    disabledReason: stats.numericCount < 2 ? 'requiresTwoNumbers' : undefined,
  });
  out.push({
    id: 'format-icon-set',
    group: 'formatting',
    labelKey: 'iconSet',
    disabled: stats.numericCount < 3,
    disabledReason: stats.numericCount < 3 ? 'requiresThreeNumbers' : undefined,
  });
  out.push({
    id: 'format-greater-than',
    group: 'formatting',
    labelKey: 'greaterThan',
    disabled: !hasNumbers,
    disabledReason: !hasNumbers ? 'requiresNumbers' : undefined,
  });
  out.push({
    id: 'format-top-10',
    group: 'formatting',
    labelKey: 'top10',
    disabled: stats.numericCount < 2,
    disabledReason: stats.numericCount < 2 ? 'requiresTwoNumbers' : undefined,
  });
  out.push({ id: 'format-clear', group: 'formatting', labelKey: 'clearFormat' });

  // Totals section.
  out.push({
    id: 'totals-sum-row',
    group: 'totals',
    labelKey: 'sumRow',
    disabled: !hasNumbers || !multi,
    disabledReason: totalsReason,
  });
  out.push({
    id: 'totals-sum-col',
    group: 'totals',
    labelKey: 'sumCol',
    disabled: !hasNumbers || !multi,
    disabledReason: totalsReason,
  });
  out.push({
    id: 'totals-average-row',
    group: 'totals',
    labelKey: 'avgRow',
    disabled: !hasNumbers || !multi,
    disabledReason: totalsReason,
  });
  out.push({
    id: 'totals-count-row',
    group: 'totals',
    labelKey: 'countRow',
    disabled: !multi,
    disabledReason: !multi ? 'requiresMultiCell' : undefined,
  });

  // Tables section.
  out.push({
    id: 'tables-as-table',
    group: 'tables',
    labelKey: 'formatAsTable',
    disabled: !multi,
    disabledReason: !multi ? 'requiresMultiCell' : undefined,
  });
  out.push({
    id: 'tables-pivot',
    group: 'tables',
    labelKey: 'pivotTable',
    disabled: !multi || input.pivotTableAvailable !== true,
    disabledReason: !multi
      ? 'requiresMultiCell'
      : input.pivotTableAvailable !== true
        ? 'pivotUnavailable'
        : undefined,
  });

  // Sparkline section — only meaningful for a horizontal series.
  const horizontalRun = range.r0 === range.r1 && range.c1 > range.c0;
  out.push({
    id: 'sparkline-line',
    group: 'sparklines',
    labelKey: 'sparkLine',
    disabled: !horizontalRun,
    disabledReason: !horizontalRun ? 'requiresHorizontalRun' : undefined,
  });
  out.push({
    id: 'sparkline-column',
    group: 'sparklines',
    labelKey: 'sparkColumn',
    disabled: !horizontalRun,
    disabledReason: !horizontalRun ? 'requiresHorizontalRun' : undefined,
  });
  out.push({
    id: 'sparkline-win-loss',
    group: 'sparklines',
    labelKey: 'sparkWinLoss',
    disabled: !horizontalRun,
    disabledReason: !horizontalRun ? 'requiresHorizontalRun' : undefined,
  });

  out.push({
    id: 'charts-column',
    group: 'charts',
    labelKey: 'chartColumn',
    disabled: !hasNumbers || !multi || input.chartAvailable !== true,
    disabledReason: chartReason,
  });
  out.push({
    id: 'charts-line',
    group: 'charts',
    labelKey: 'chartLine',
    disabled: !hasNumbers || !multi || input.chartAvailable !== true,
    disabledReason: chartReason,
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

function firstLockedProtectedAddr(input: QuickAnalysisExecuteInput, range: Range) {
  const state = input.store.getState();
  if (!isSheetProtected(state, range.sheet)) return null;
  if (state.protection.allowedEditRanges.some((entry) => rangeContainsRange(entry.range, range))) {
    return null;
  }
  if (rangeArea(range) > MAX_EXACT_PROTECTION_SCAN_CELLS) {
    return { sheet: range.sheet, row: range.r0, col: range.c0 };
  }
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const addr = { sheet: range.sheet, row, col };
      if (!isCellWritable(state, addr)) return addr;
    }
  }
  return null;
}

function addConditionalFormat(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { actionId, range, stats, store } = input;
  const locked = firstLockedProtectedAddr(input, range);
  if (locked) {
    warnProtected(locked);
    return { ok: false, reason: 'protected' };
  }
  const add = (rule: Parameters<typeof mutators.addConditionalRule>[1]) => {
    recordConditionalRulesChange(input.history ?? null, store, () => {
      mutators.addConditionalRule(store, rule);
    });
  };
  if (actionId === 'format-data-bar') {
    add({
      kind: 'data-bar',
      range,
      color: '#5b9bd5',
      showValue: true,
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-color-scale') {
    add({
      kind: 'color-scale',
      range,
      stops: ['#f8696b', '#ffeb84', '#63be7b'],
      thresholds: [{ kind: 'min' }, { kind: 'percentile', value: 50 }, { kind: 'max' }],
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-icon-set') {
    add({
      kind: 'icon-set',
      range,
      icons: 'traffic3',
      showValue: true,
      thresholds: [
        { kind: 'percent', value: 100 / 3 },
        { kind: 'percent', value: 200 / 3 },
      ],
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-greater-than') {
    add({
      kind: 'cell-value',
      range,
      op: '>',
      a: stats.avg,
      apply: { fill: '#ffc7ce', color: '#9c0006' },
    });
    return { ok: true, kind: 'conditional-format', count: 1 };
  }
  if (actionId === 'format-top-10') {
    add({
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
  const { actionId, range, store, wb } = input;
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
    if (range.c1 - range.c0 + 1 > MAX_QUICK_ANALYSIS_FORMULA_WRITES) {
      return { ok: false, reason: 'out-of-bounds' };
    }
    const writes: Array<{ addr: { sheet: number; row: number; col: number }; formula: string }> =
      [];
    for (let col = range.c0; col <= range.c1; col += 1) {
      const addr = { sheet: range.sheet, row, col };
      if (!isCellWritable(store.getState(), addr)) {
        warnProtected(addr);
        continue;
      }
      const formula = `=${fn}(${colLetter(col)}${range.r0 + 1}:${colLetter(col)}${range.r1 + 1})`;
      writes.push({ addr, formula });
    }
    if (writes.length === 0) return { ok: false, reason: 'protected' };
    if (input.history) input.history.begin();
    try {
      for (const write of writes) {
        wb.setFormula(write.addr, write.formula);
        count += 1;
      }
    } finally {
      if (input.history) input.history.end();
    }
    return { ok: true, kind: 'formula', count };
  }

  if (actionId === 'totals-sum-col') {
    const col = range.c1 + 1;
    if (col > MAX_COL) return { ok: false, reason: 'out-of-bounds' };
    if (range.r1 - range.r0 + 1 > MAX_QUICK_ANALYSIS_FORMULA_WRITES) {
      return { ok: false, reason: 'out-of-bounds' };
    }
    const writes: Array<{ addr: { sheet: number; row: number; col: number }; formula: string }> =
      [];
    for (let row = range.r0; row <= range.r1; row += 1) {
      const addr = { sheet: range.sheet, row, col };
      if (!isCellWritable(store.getState(), addr)) {
        warnProtected(addr);
        continue;
      }
      const formula = `=SUM(${colLetter(range.c0)}${row + 1}:${colLetter(range.c1)}${row + 1})`;
      writes.push({ addr, formula });
    }
    if (writes.length === 0) return { ok: false, reason: 'protected' };
    if (input.history) input.history.begin();
    try {
      for (const write of writes) {
        wb.setFormula(write.addr, write.formula);
        count += 1;
      }
    } finally {
      if (input.history) input.history.end();
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
  const addr = { sheet: range.sheet, row: range.r0, col };
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return { ok: false, reason: 'protected' };
  }
  recordSparklineChange(input.history ?? null, store, () => {
    mutators.setSparkline(store, addr, {
      kind,
      source: rangeRef(range),
      showNegative: kind !== 'line',
    });
  });
  return { ok: true, kind: 'sparkline', count: 1 };
}

function addChart(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { actionId, range, store } = input;
  if (isSheetProtected(store.getState(), range.sheet)) return { ok: false, reason: 'protected' };
  const kind = actionId === 'charts-line' ? 'line' : 'column';
  const count = store.getState().charts.charts.length;
  recordChartsChange(input.history ?? null, store, () => {
    createSessionChart(store, range, {
      id: `qa-chart-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}-${kind}`,
      kind,
      title: null,
      x: 320 + (count % 3) * 24,
      y: 72 + (count % 3) * 24,
      w: 360,
      h: 220,
    });
  });
  return { ok: true, kind: 'chart', count: 1 };
}

function formatAsTable(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { range, store } = input;
  let ok = false;
  recordTablesChange(input.history ?? null, store, () => {
    ok =
      applyFormatAsTable(store, range, {
        id: `qa-table-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}`,
      }) != null;
  });
  if (!ok) return { ok: false, reason: 'protected' };
  return { ok: true, kind: 'table', count: 1 };
}

function clearAnalysisFormatting(input: QuickAnalysisExecuteInput): QuickAnalysisExecuteResult {
  const { store, range } = input;
  const history = input.history ?? null;
  const locked = firstLockedProtectedAddr(input, range);
  if (locked) {
    warnProtected(locked);
    return { ok: false, reason: 'protected' };
  }
  if (history) history.begin();
  try {
    recordFormatChange(history, store, () => {
      store.setState((s) => {
        const formats = new Map(s.format.formats);
        for (const key of s.format.formats.keys()) {
          const addr = addrFromKey(key);
          if (!addr || !rangeContainsAddr(range, addr)) continue;
          if (!isCellWritable(s, addr)) continue;
          formats.delete(key);
        }
        return { ...s, format: { ...s.format, formats } };
      });
    });
    if (!isSheetProtected(store.getState(), range.sheet)) {
      recordConditionalRulesChange(history, store, () => {
        mutators.clearConditionalRulesInRange(store, range);
      });
      recordSparklineChange(history, store, () => {
        mutators.clearSparklinesInRange(store, range);
      });
      recordChartsChange(history, store, () => {
        mutators.clearChartsInRange(store, range);
      });
      recordTablesChange(history, store, () => {
        mutators.clearTableOverlaysInRange(store, range);
      });
    }
  } finally {
    if (history) history.end();
  }
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
