import { describe, expect, it } from 'vitest';
import {
  applyPivotEdit,
  type PivotEdit,
  type PivotEditField,
} from '../../../src/commands/pivot-edit.js';
import {
  PivotAggregation,
  PivotAxis,
  type PivotFilterSpec,
  PivotReportLayout,
} from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const pivot = { sheetIndex: 1, pivotIndex: 2, rows: 5, cols: 4, fields: ['Region', 'Sales'] };

const field = (over: Partial<PivotEditField> & { fieldIndex: number }): PivotEditField => ({
  axis: PivotAxis.Row,
  aggregation: PivotAggregation.Sum,
  numberFormat: '',
  filterItems: [],
  filterChecklist: null,
  ...over,
});

const edit = (over: Partial<PivotEdit> = {}): PivotEdit => ({
  fieldListOnly: false,
  name: 'Sales pivot',
  anchor: { row: 3, col: 6 },
  rowGrandTotals: true,
  colGrandTotals: false,
  layout: PivotReportLayout.Tabular,
  fields: [],
  filters: null,
  ...over,
});

const fakeWorkbook = (
  opts: { failing?: string[]; dataFieldCount?: number; byIndex?: boolean } = {},
) => {
  const calls: unknown[][] = [];
  const failing = new Set(opts.failing ?? []);
  const record =
    (name: string, ok: boolean | number = true) =>
    (...args: unknown[]) => {
      calls.push([name, ...args]);
      return failing.has(name) ? false : ok;
    };
  const wb = {
    renamePivotTable: record('rename'),
    setPivotTableAnchor: record('anchor'),
    setPivotTableGrandTotals: record('totals'),
    setPivotReportLayout: record('layout'),
    setPivotFieldAxis: record('axis'),
    setPivotRowFieldOrder: record('rowOrder'),
    setPivotColFieldOrder: record('colOrder'),
    pivotDataFieldCount: () => opts.dataFieldCount ?? 0,
    setPivotDataField: record('setData'),
    addPivotDataField: (...args: unknown[]) => {
      calls.push(['addData', ...args]);
      return failing.has('addData') ? -1 : 0;
    },
    clearPivotFieldItems: record('clearItems'),
    addPivotFieldItemAt: record('itemAt', opts.byIndex ?? true),
    addPivotFieldItem: record('item'),
    clearPivotFilters: record('clearFilters'),
    addPivotFilter: record('addFilter'),
  } as unknown as WorkbookHandle;
  return { wb, calls, names: () => calls.map((c) => c[0]) };
};

describe('applyPivotEdit', () => {
  it('writes the header settings and field axes, ordering row and column fields', () => {
    const { wb, calls } = fakeWorkbook();
    const ok = applyPivotEdit(
      wb,
      pivot,
      edit({
        fields: [
          field({ fieldIndex: 0, axis: PivotAxis.Col }),
          field({ fieldIndex: 1, axis: PivotAxis.Row }),
        ],
      }),
    );
    expect(ok).toBe(true);
    expect(calls).toEqual([
      ['rename', 1, 2, 'Sales pivot'],
      ['anchor', 1, 2, { row: 3, col: 6, rows: 5, cols: 4 }],
      ['totals', 1, 2, true, false],
      ['layout', 1, 2, PivotReportLayout.Tabular],
      ['axis', 1, 2, 0, PivotAxis.Col],
      ['axis', 1, 2, 1, PivotAxis.Row],
      ['rowOrder', 1, 2, [1]],
      ['colOrder', 1, 2, [0]],
    ]);
  });

  it('skips the header writes for a field-list edit', () => {
    const { wb, names } = fakeWorkbook();
    applyPivotEdit(wb, pivot, edit({ fieldListOnly: true, fields: [field({ fieldIndex: 0 })] }));
    expect(names()).toEqual(['axis', 'rowOrder', 'colOrder']);
  });

  it('keeps writing after a failed step but reports failure', () => {
    const { wb, names } = fakeWorkbook({ failing: ['rename'] });
    const ok = applyPivotEdit(wb, pivot, edit({ fields: [field({ fieldIndex: 0 })] }));
    expect(ok).toBe(false);
    expect(names()).toContain('layout');
    expect(names()).toContain('rowOrder');
  });

  it('skips axis ordering when an axis write fails', () => {
    const { wb, names } = fakeWorkbook({ failing: ['axis'] });
    expect(applyPivotEdit(wb, pivot, edit({ fields: [field({ fieldIndex: 0 })] }))).toBe(false);
    expect(names()).not.toContain('rowOrder');
  });

  it('updates existing data fields in place and adds the rest', () => {
    const { wb, calls } = fakeWorkbook({ dataFieldCount: 1 });
    const ok = applyPivotEdit(
      wb,
      pivot,
      edit({
        fieldListOnly: true,
        fields: [
          field({ fieldIndex: 0, axis: PivotAxis.Value, aggregation: PivotAggregation.Count }),
          field({
            fieldIndex: 1,
            axis: PivotAxis.Value,
            aggregation: PivotAggregation.Sum,
            numberFormat: '0.00',
          }),
        ],
      }),
    );
    expect(ok).toBe(true);
    const data = calls.filter(([name]) => name === 'setData' || name === 'addData');
    expect(data).toHaveLength(2);
    expect(data[0]?.[0]).toBe('setData');
    expect(data[0]?.[3]).toBe(0);
    expect(data[0]?.[4]).toMatchObject({ fieldIndex: 0, aggregation: PivotAggregation.Count });
    expect(data[0]?.[4]).not.toHaveProperty('numberFormat');
    expect(data[1]?.[0]).toBe('addData');
    expect(data[1]?.[3]).toMatchObject({ fieldIndex: 1, numberFormat: '0.00' });
  });

  it('writes checklist items by cache index and falls back to the label', () => {
    const { wb, calls } = fakeWorkbook({ byIndex: false });
    const ok = applyPivotEdit(
      wb,
      pivot,
      edit({
        fieldListOnly: true,
        fields: [
          field({
            fieldIndex: 0,
            axis: PivotAxis.Page,
            filterChecklist: [
              { value: 'East', checked: true, cacheIndex: 0 },
              { value: '', checked: true, cacheIndex: 1 },
              { value: 'West', checked: false },
            ],
          }),
        ],
      }),
    );
    expect(ok).toBe(true);
    expect(calls.filter(([name]) => name === 'item')).toEqual([
      ['item', 1, 2, 0, 'East', true],
      ['item', 1, 2, 0, 'West', false],
    ]);
  });

  it('adds typed filter items as visible when no checklist is known', () => {
    const { wb, calls } = fakeWorkbook();
    applyPivotEdit(
      wb,
      pivot,
      edit({
        fieldListOnly: true,
        fields: [field({ fieldIndex: 1, axis: PivotAxis.Page, filterItems: ['a', 'b'] })],
      }),
    );
    expect(calls.filter(([name]) => name === 'clearItems' || name === 'item')).toEqual([
      ['clearItems', 1, 2, 1],
      ['item', 1, 2, 1, 'a', true],
      ['item', 1, 2, 1, 'b', true],
    ]);
  });

  it('replaces filters only when a replacement is given', () => {
    const spec = { fieldName: 'Region' } as PivotFilterSpec;
    const untouched = fakeWorkbook();
    applyPivotEdit(untouched.wb, pivot, edit({ fieldListOnly: true, filters: null }));
    expect(untouched.names()).not.toContain('clearFilters');

    const replaced = fakeWorkbook();
    const ok = applyPivotEdit(replaced.wb, pivot, edit({ fieldListOnly: true, filters: [spec] }));
    expect(ok).toBe(true);
    expect(replaced.calls.slice(-2)).toEqual([
      ['clearFilters', 1, 2],
      ['addFilter', 1, 2, spec],
    ]);
  });
});
