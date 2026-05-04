import { describe, expect, it } from 'vitest';
import type { SelectionStats } from '../../../src/commands/aggregate.js';
import {
  buildQuickAnalysisActions,
  groupQuickAnalysisActions,
} from '../../../src/commands/quick-analysis.js';
import type { Range } from '../../../src/engine/types.js';

const mkStats = (overrides: Partial<SelectionStats> = {}): SelectionStats => ({
  cells: 1,
  numericCount: 0,
  nonBlankCount: 0,
  sum: 0,
  avg: 0,
  min: 0,
  max: 0,
  ...overrides,
});

const range = (r0: number, c0: number, r1: number, c1: number): Range => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});

const findAction = (id: string, list: ReturnType<typeof buildQuickAnalysisActions>) =>
  list.find((a) => a.id === id);

describe('buildQuickAnalysisActions', () => {
  it('emits actions for every group', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 0, 0),
      stats: mkStats(),
    });
    const groups = new Set(actions.map((a) => a.group));
    expect(groups).toEqual(new Set(['formatting', 'totals', 'tables', 'sparklines', 'charts']));
  });

  it('disables data-bar / color-scale on a single numeric cell', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 0, 0),
      stats: mkStats({ numericCount: 1 }),
    });
    expect(findAction('format-data-bar', actions)?.disabled).toBe(true);
    expect(findAction('format-color-scale', actions)?.disabled).toBe(true);
  });

  it('enables data-bar when the range has at least two numbers', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 4, 0),
      stats: mkStats({ numericCount: 2 }),
    });
    expect(findAction('format-data-bar', actions)?.disabled).toBe(false);
  });

  it('disables totals on a single-cell selection even with numbers', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 0, 0),
      stats: mkStats({ numericCount: 1 }),
    });
    expect(findAction('totals-sum-row', actions)?.disabled).toBe(true);
  });

  it('enables totals on a multi-cell numeric range', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 4, 4),
      stats: mkStats({ numericCount: 8 }),
    });
    expect(findAction('totals-sum-row', actions)?.disabled).toBe(false);
    expect(findAction('totals-sum-col', actions)?.disabled).toBe(false);
  });

  it('only enables sparkline actions on a horizontal run (single row)', () => {
    const horizontal = buildQuickAnalysisActions({
      range: range(0, 0, 0, 4),
      stats: mkStats({ numericCount: 5 }),
    });
    expect(findAction('sparkline-line', horizontal)?.disabled).toBe(false);

    const block = buildQuickAnalysisActions({
      range: range(0, 0, 4, 4),
      stats: mkStats({ numericCount: 5 }),
    });
    expect(findAction('sparkline-line', block)?.disabled).toBe(true);
  });

  it('disables Format As Table on a single cell', () => {
    const single = buildQuickAnalysisActions({
      range: range(0, 0, 0, 0),
      stats: mkStats({ numericCount: 1 }),
    });
    expect(findAction('tables-as-table', single)?.disabled).toBe(true);

    const multi = buildQuickAnalysisActions({
      range: range(0, 0, 4, 4),
      stats: mkStats({ numericCount: 1 }),
    });
    expect(findAction('tables-as-table', multi)?.disabled).toBe(false);
  });

  it('keeps charts + pivot stubs disabled (engine integration pending)', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 5, 5),
      stats: mkStats({ numericCount: 25 }),
    });
    expect(findAction('charts-placeholder', actions)?.disabled).toBe(true);
    expect(findAction('tables-pivot', actions)?.disabled).toBe(true);
  });
});

describe('groupQuickAnalysisActions', () => {
  it('partitions actions by group preserving order', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 4, 4),
      stats: mkStats({ numericCount: 5 }),
    });
    const grouped = groupQuickAnalysisActions(actions);
    expect(grouped.formatting.length).toBeGreaterThan(0);
    expect(grouped.totals.length).toBeGreaterThan(0);
    expect(grouped.tables.length).toBeGreaterThan(0);
    expect(grouped.sparklines.length).toBeGreaterThan(0);
    expect(grouped.charts.length).toBeGreaterThan(0);
    // The first formatting action is data-bar (canonical order).
    expect(grouped.formatting[0]?.id).toBe('format-data-bar');
  });
});
