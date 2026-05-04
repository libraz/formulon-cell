import type { Range } from '../engine/types.js';
import type { SelectionStats } from './aggregate.js';

/**
 * Quick Analysis (Excel: Ctrl+Q) — a context-aware action sheet that
 * appears anchored at the selection's bottom-right corner. The host
 * renders the popover; this module provides the pure logic that decides
 * which actions are relevant for a given selection.
 *
 * Sections mirror Excel:
 *   - 書式設定 (Formatting)  — fill / data-bar / icon-set tied to numbers.
 *   - グラフ (Charts)        — sparkline placeholder; real charts wait on engine.
 *   - 合計 (Totals)          — sum / average etc. for numeric ranges.
 *   - テーブル (Tables)      — Format-As-Table / Pivot stub.
 *   - スパークライン         — line / column / win-loss for numeric ranges.
 */
export type QuickAnalysisGroup = 'formatting' | 'charts' | 'totals' | 'tables' | 'sparklines';

export interface QuickAnalysisAction {
  id: string;
  group: QuickAnalysisGroup;
  /** Untranslated key — host resolves through i18n strings. */
  labelKey: string;
  /** True when the action is a no-op for the current selection. */
  disabled?: boolean;
}

export interface QuickAnalysisInput {
  /** Range under the popover anchor — typically the primary selection. */
  range: Range;
  /** Aggregate stats over the same range. Drives "totals" availability. */
  stats: SelectionStats;
}

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
  out.push({ id: 'tables-pivot', group: 'tables', labelKey: 'pivotStub', disabled: true });

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

  // Charts section — engine doesn't expose chart embind yet, so we surface
  // a single placeholder that opens a "coming soon" notice. Disabled keeps
  // the surface area honest in the meantime.
  out.push({ id: 'charts-placeholder', group: 'charts', labelKey: 'chartsStub', disabled: true });

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
