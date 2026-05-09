import { describe, expect, it } from 'vitest';
import type { SelectionStats } from '../../../src/commands/aggregate.js';
import {
  buildQuickAnalysisActions,
  enabledQuickAnalysisActions,
  executeQuickAnalysisAction,
  groupQuickAnalysisActions,
  isQuickAnalysisActionEnabled,
  quickAnalysisActionById,
} from '../../../src/commands/quick-analysis.js';
import type { Range } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

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

  it('enables session charts for multi-cell numeric ranges and pivot only when the host can open it', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 5, 5),
      stats: mkStats({ numericCount: 25 }),
      chartAvailable: true,
    });
    expect(findAction('charts-column', actions)?.disabled).toBe(false);
    expect(findAction('charts-line', actions)?.disabled).toBe(false);
    expect(findAction('tables-pivot', actions)?.disabled).toBe(true);

    const withPivot = buildQuickAnalysisActions({
      range: range(0, 0, 5, 5),
      stats: mkStats({ numericCount: 25 }),
      pivotTableAvailable: true,
      chartAvailable: true,
    });
    expect(findAction('tables-pivot', withPivot)?.disabled).toBe(false);
  });

  it('disables chart actions when the host chart renderer is unavailable', () => {
    const actions = buildQuickAnalysisActions({
      range: range(0, 0, 5, 5),
      stats: mkStats({ numericCount: 25 }),
      chartAvailable: false,
    });
    expect(findAction('charts-column', actions)?.disabled).toBe(true);
    expect(findAction('charts-line', actions)?.disabled).toBe(true);
  });

  it('exposes action lookup and enabled-state helpers for custom hosts', () => {
    const input = {
      range: range(0, 0, 3, 3),
      stats: mkStats({ numericCount: 8 }),
      chartAvailable: true,
    };
    const actions = buildQuickAnalysisActions(input);

    expect(quickAnalysisActionById(actions, 'charts-column')).toMatchObject({
      id: 'charts-column',
      disabled: false,
    });
    expect(quickAnalysisActionById(actions, 'tables-pivot')?.disabled).toBe(true);
    expect(isQuickAnalysisActionEnabled(input, 'charts-column')).toBe(true);
    expect(isQuickAnalysisActionEnabled(input, 'tables-pivot')).toBe(false);
    expect(enabledQuickAnalysisActions(input).some((action) => action.id === 'format-clear')).toBe(
      true,
    );
    expect(enabledQuickAnalysisActions(input).some((action) => action.id === 'tables-pivot')).toBe(
      false,
    );
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

describe('executeQuickAnalysisAction', () => {
  const makeWb = (): {
    wb: WorkbookHandle;
    formulas: Array<{ addr: { sheet: number; row: number; col: number }; formula: string }>;
  } => {
    const formulas: Array<{ addr: { sheet: number; row: number; col: number }; formula: string }> =
      [];
    return {
      wb: {
        setFormula(addr: { sheet: number; row: number; col: number }, formula: string) {
          formulas.push({ addr, formula });
          return true;
        },
      } as unknown as WorkbookHandle,
      formulas,
    };
  };

  it('adds conditional formatting rules for enabled formatting actions', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'format-data-bar',
      range: range(0, 0, 4, 0),
      stats: mkStats({ numericCount: 5 }),
    });

    expect(result).toEqual({ ok: true, kind: 'conditional-format', count: 1 });
    expect(store.getState().conditional.rules).toEqual([
      {
        kind: 'data-bar',
        range: range(0, 0, 4, 0),
        color: '#5b9bd5',
        showValue: true,
      },
    ]);
  });

  it('writes row total formulas below the selected range', () => {
    const store = createSpreadsheetStore();
    const { wb, formulas } = makeWb();
    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'totals-sum-row',
      range: range(1, 0, 3, 1),
      stats: mkStats({ numericCount: 6 }),
    });

    expect(result).toEqual({ ok: true, kind: 'formula', count: 2 });
    expect(formulas).toEqual([
      { addr: { sheet: 0, row: 4, col: 0 }, formula: '=SUM(A2:A4)' },
      { addr: { sheet: 0, row: 4, col: 1 }, formula: '=SUM(B2:B4)' },
    ]);
  });

  it('writes column total formulas to the right of the selected range', () => {
    const store = createSpreadsheetStore();
    const { wb, formulas } = makeWb();
    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'totals-sum-col',
      range: range(0, 1, 1, 3),
      stats: mkStats({ numericCount: 6 }),
    });

    expect(result).toEqual({ ok: true, kind: 'formula', count: 2 });
    expect(formulas).toEqual([
      { addr: { sheet: 0, row: 0, col: 4 }, formula: '=SUM(B1:D1)' },
      { addr: { sheet: 0, row: 1, col: 4 }, formula: '=SUM(B2:D2)' },
    ]);
  });

  it('adds a sparkline in the adjacent cell for horizontal ranges', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'sparkline-column',
      range: range(2, 0, 2, 3),
      stats: mkStats({ numericCount: 4 }),
    });

    expect(result).toEqual({ ok: true, kind: 'sparkline', count: 1 });
    expect(store.getState().sparkline.sparklines.get('0:2:4')).toEqual({
      kind: 'column',
      source: 'A3:D3',
      showNegative: true,
    });
  });

  it('adds a session chart overlay for chart actions', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    const r = range(0, 0, 4, 2);
    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'charts-line',
      range: r,
      stats: mkStats({ numericCount: 10 }),
      chartAvailable: true,
    });

    expect(result).toEqual({ ok: true, kind: 'chart', count: 1 });
    expect(store.getState().charts.charts).toEqual([
      expect.objectContaining({
        id: 'qa-chart-0-0-0-4-2-line',
        kind: 'line',
        source: r,
      }),
    ]);
    expect(store.getState().charts.charts[0]?.title).toBeUndefined();
  });

  it('adds a session table overlay for Format As Table', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    const r = range(0, 0, 4, 2);
    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'tables-as-table',
      range: r,
      stats: mkStats({ numericCount: 6 }),
    });

    expect(result).toEqual({ ok: true, kind: 'table', count: 1 });
    expect(store.getState().tables.tables).toEqual([
      {
        id: 'qa-table-0-0-0-4-2',
        source: 'session',
        range: r,
        style: 'medium',
        showHeader: true,
        showTotal: false,
        banded: true,
      },
    ]);
  });

  it('clears static formats and Quick Analysis overlays in the selected range', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    const r = range(0, 0, 4, 2);
    mutators.setRange(store, r);
    mutators.setRangeFormat(store, r, { fill: '#ffff00' });
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: r,
      color: '#5b9bd5',
      showValue: true,
    });
    mutators.setSparkline(store, { sheet: 0, row: 1, col: 1 }, { kind: 'line', source: 'A1:C1' });
    mutators.upsertChart(store, {
      id: 'chart',
      kind: 'column',
      source: r,
      title: 'Column chart',
    });
    mutators.upsertTableOverlay(store, {
      id: 'tbl',
      source: 'session',
      range: r,
      style: 'medium',
      showHeader: true,
      showTotal: false,
      banded: true,
    });

    const result = executeQuickAnalysisAction({
      store,
      wb,
      actionId: 'format-clear',
      range: r,
      stats: mkStats({ numericCount: 6 }),
    });

    expect(result).toEqual({ ok: true, kind: 'clear-format', count: 1 });
    expect(store.getState().format.formats.size).toBe(0);
    expect(store.getState().conditional.rules).toHaveLength(0);
    expect(store.getState().sparkline.sparklines.size).toBe(0);
    expect(store.getState().charts.charts).toHaveLength(0);
    expect(store.getState().tables.tables).toHaveLength(0);
  });

  it('does not execute disabled or unsupported actions', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();

    expect(
      executeQuickAnalysisAction({
        store,
        wb,
        actionId: 'sparkline-line',
        range: range(0, 0, 3, 0),
        stats: mkStats({ numericCount: 4 }),
      }),
    ).toEqual({ ok: false, reason: 'disabled' });
    expect(
      executeQuickAnalysisAction({
        store,
        wb,
        actionId: 'charts-column',
        range: range(0, 0, 0, 0),
        stats: mkStats({ numericCount: 1 }),
        chartAvailable: true,
      }),
    ).toEqual({ ok: false, reason: 'disabled' });
  });
});
