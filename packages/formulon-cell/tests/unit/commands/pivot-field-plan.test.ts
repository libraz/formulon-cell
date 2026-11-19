import { describe, expect, it } from 'vitest';
import {
  type CreatePivotTableOptions,
  createPivotTableFromRange,
} from '../../../src/commands/pivot-table.js';
import {
  PivotAggregation,
  PivotAxis,
  PivotFilterKind,
  PivotShowValuesAs,
} from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const SOURCE = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 3 };
const DEST = { sheet: 0, row: 8, col: 0 };

// Region, Product, Channel are text; Amt is numeric.
const setup = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Channel', 'Amt'],
    ['East', 'A', 'Web', 500],
    ['North', 'A', 'Shop', 100],
    ['North', 'B', 'Web', 200],
    ['South', 'C', 'Shop', 300],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  return wb;
};

const create = (wb: WorkbookHandle, over: Partial<CreatePivotTableOptions>) =>
  createPivotTableFromRange(wb, {
    source: SOURCE,
    destination: DEST,
    rowField: 'Region',
    valueField: 'Amt',
    ...over,
  });

/** Source name of each field on the given axis, in pivot field order. */
const onAxis = (wb: WorkbookHandle, axis: PivotAxis): string[] =>
  (wb.getPivotDefinition(0, 0)?.fields ?? [])
    .filter((f) => f.axis === axis)
    .map((f) => f.sourceName);

const dataFields = (wb: WorkbookHandle) => {
  const def = wb.getPivotDefinition(0, 0);
  return (def?.dataFields ?? []).map((d) => ({
    source: def?.fields[d.fieldIndex]?.sourceName,
    aggregation: d.aggregation,
    numberFormat: d.numberFormat ?? '',
    showValuesAs: d.showValuesAs,
  }));
};

describe('createPivotTableFromRange field refusals', () => {
  const refusals: [string, Partial<CreatePivotTableOptions>][] = [
    [
      'filterItems naming a field that is not a filter field',
      {
        filterField: 'Channel',
        filterItems: [{ fieldName: 'Product', itemName: 'A', visible: false }],
      },
    ],
    ['a filter field equal to the row field', { filterField: 'Region' }],
    [
      'a filter field equal to the column field',
      { columnField: 'Product', filterField: 'Product' },
    ],
    ['a plural filter list containing the row field', { filterFields: ['Channel', 'Region'] }],
    ['row field equal to column field', { columnField: 'Region' }],
    ['a value field that does not exist', { valueField: 'Nope' }],
    ['a missing name among several value fields', { valueFields: ['Amt', 'Nope'] }],
    ['a filter field that does not exist', { filterField: 'Nope' }],
    ['an unknown columnField', { columnField: 'Nope' }],
    ['an unknown rowField', { rowField: 'Nope' }],
    ['a row field with no value field', { valueField: '' }],
    ['values with no row field', { rowField: '' }],
    [
      'a pivotFilters entry on a value field',
      {
        pivotFilters: [
          {
            axis: PivotAxis.Row,
            fieldName: 'Amt',
            kind: PivotFilterKind.LabelContains,
            value: '1',
          },
        ],
      },
    ],
    [
      'showAsBaseField that is not a row or column field',
      {
        valueFieldSettings: [
          {
            fieldName: 'Amt',
            showValuesAs: PivotShowValuesAs.PercentOf,
            showAsBaseField: 'Amt',
          },
        ],
      },
    ],
    [
      'showAsBaseField that is a filter field',
      {
        filterField: 'Channel',
        valueFieldSettings: [
          {
            fieldName: 'Amt',
            showValuesAs: PivotShowValuesAs.DifferenceFrom,
            showAsBaseField: 'Channel',
          },
        ],
      },
    ],
  ];

  it.each(refusals)('refuses %s and removes the cache it built', async (_name, over) => {
    const wb = await setup();
    try {
      const before = wb.pivotCacheCount();
      expect(create(wb, over)).toEqual({ ok: false, reason: 'invalid-field' });
      expect(wb.pivotCacheCount()).toBe(before);
      expect(wb.getPivotDefinition(0, 0)).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('keeps a shared cache when it refuses a field', async () => {
    const wb = await setup();
    try {
      const first = create(wb, {});
      expect(first.ok).toBe(true);
      if (!first.ok) return;
      const ids = wb.pivotCacheIds();
      const refused = create(wb, { cacheId: first.cacheId, valueField: 'Nope' });
      expect(refused).toEqual({ ok: false, reason: 'invalid-field' });
      expect(wb.pivotCacheIds()).toEqual(ids);
      expect(wb.getPivotTables()).toHaveLength(1);
    } finally {
      wb.dispose();
    }
  });

  it('refuses a source with fewer than two fields as an invalid range and removes the cache', async () => {
    const wb = await setup();
    try {
      const before = wb.pivotCacheCount();
      const result = create(wb, { source: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 } });
      expect(result).toEqual({ ok: false, reason: 'invalid-range' });
      expect(wb.pivotCacheCount()).toBe(before);
    } finally {
      wb.dispose();
    }
  });

  it('keeps a shared cache when the source has fewer than two fields', async () => {
    const wb = await setup();
    try {
      const narrow = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
      const cacheId = wb.createPivotCache(0);
      expect(cacheId).toBeGreaterThanOrEqual(0);
      // An unrefreshed cache has no fields, so it is below the two-field minimum.
      const result = create(wb, { source: narrow, cacheId });
      expect(result).toEqual({ ok: false, reason: 'invalid-range' });
      expect(wb.pivotCacheIds()).toContain(cacheId);
    } finally {
      wb.dispose();
    }
  });
});

describe('createPivotTableFromRange field roles', () => {
  it('creates an empty report with every field off its axis', async () => {
    const wb = await setup();
    try {
      const result = create(wb, { rowField: '', valueField: '' });
      expect(result).toMatchObject({ ok: true, sharedCache: false });
      const def = wb.getPivotDefinition(0, 0);
      expect(def?.fields.map((f) => f.sourceName)).toEqual(['Region', 'Product', 'Channel', 'Amt']);
      expect(def?.fields.every((f) => f.axis === null)).toBe(true);
      expect(def?.dataFields).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('creates a layout-only report from a row and a column field', async () => {
    const wb = await setup();
    try {
      const result = create(wb, { columnField: 'Product', valueField: '' });
      expect(result.ok).toBe(true);
      expect(onAxis(wb, PivotAxis.Row)).toEqual(['Region']);
      expect(onAxis(wb, PivotAxis.Col)).toEqual(['Product']);
      expect(dataFields(wb)).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('maps row, column, filter and value fields to their axes', async () => {
    const wb = await setup();
    try {
      const result = create(wb, { columnField: 'Product', filterField: 'Channel' });
      expect(result.ok).toBe(true);
      expect(onAxis(wb, PivotAxis.Row)).toEqual(['Region']);
      expect(onAxis(wb, PivotAxis.Col)).toEqual(['Product']);
      expect(onAxis(wb, PivotAxis.Page)).toEqual(['Channel']);
      expect(dataFields(wb).map((d) => d.source)).toEqual(['Amt']);
    } finally {
      wb.dispose();
    }
  });

  it('prefers the plural filterFields over filterField and drops duplicates', async () => {
    const wb = await setup();
    try {
      const result = create(wb, {
        filterField: 'Product',
        filterFields: ['Channel', 'Channel', 'Amt'],
        valueField: 'Region',
      });
      expect(result.ok).toBe(true);
      // Product is ignored; Amt sits on the page axis and so is not a value role.
      expect(onAxis(wb, PivotAxis.Page).sort()).toEqual(['Amt', 'Channel']);
    } finally {
      wb.dispose();
    }
  });

  it('applies a repeated filter field once, so its items are listed once', async () => {
    const wb = await setup();
    try {
      const result = create(wb, {
        filterFields: ['Channel', 'Channel'],
        filterItems: [{ fieldName: 'Channel', itemName: 'Web', visible: false }],
      });
      expect(result.ok).toBe(true);
      const channel = wb.getPivotDefinition(0, 0)?.fields.find((f) => f.sourceName === 'Channel');
      expect(channel?.items.map((i) => [i.label, i.visible])).toEqual([['Web', false]]);
    } finally {
      wb.dispose();
    }
  });

  it('falls back to the singular filterField when the plural list is empty', async () => {
    const wb = await setup();
    try {
      const result = create(wb, { filterField: 'Channel', filterFields: [] });
      expect(result.ok).toBe(true);
      expect(onAxis(wb, PivotAxis.Page)).toEqual(['Channel']);
    } finally {
      wb.dispose();
    }
  });

  it('prefers the plural valueFields over valueField and drops duplicates', async () => {
    const wb = await setup();
    try {
      const result = create(wb, { valueField: 'Region', valueFields: ['Amt', 'Amt', 'Channel'] });
      expect(result.ok).toBe(true);
      expect(dataFields(wb).map((d) => d.source)).toEqual(['Amt', 'Channel']);
    } finally {
      wb.dispose();
    }
  });

  it('falls back to valueField when the plural list is empty', async () => {
    const wb = await setup();
    try {
      expect(create(wb, { valueFields: [] }).ok).toBe(true);
      expect(dataFields(wb).map((d) => d.source)).toEqual(['Amt']);
    } finally {
      wb.dispose();
    }
  });

  it('accepts a show-as base field that is the column field', async () => {
    const wb = await setup();
    try {
      const result = create(wb, {
        columnField: 'Product',
        valueFieldSettings: [
          {
            fieldName: 'Amt',
            showValuesAs: PivotShowValuesAs.PercentOf,
            showAsBaseField: 'Product',
          },
        ],
      });
      expect(result.ok).toBe(true);
    } finally {
      wb.dispose();
    }
  });
});

describe('createPivotTableFromRange value field defaults', () => {
  it('sums a numeric field and counts a text field by default', async () => {
    const wb = await setup();
    try {
      expect(create(wb, { valueFields: ['Amt', 'Channel'] }).ok).toBe(true);
      expect(dataFields(wb).map((d) => [d.source, d.aggregation])).toEqual([
        ['Amt', PivotAggregation.Sum],
        ['Channel', PivotAggregation.Count],
      ]);
    } finally {
      wb.dispose();
    }
  });

  it('lets a per-field aggregation override the report aggregation', async () => {
    const wb = await setup();
    try {
      const result = create(wb, {
        valueFields: ['Amt', 'Channel'],
        aggregation: PivotAggregation.Max,
        valueFieldSettings: [{ fieldName: 'Channel', aggregation: PivotAggregation.Count }],
      });
      expect(result.ok).toBe(true);
      expect(dataFields(wb).map((d) => [d.source, d.aggregation])).toEqual([
        ['Amt', PivotAggregation.Max],
        ['Channel', PivotAggregation.Count],
      ]);
    } finally {
      wb.dispose();
    }
  });

  const percentKinds = [
    'PercentOfRow',
    'PercentOfCol',
    'PercentOfTotal',
    'PercentOf',
    'PercentDifferenceFrom',
    'PercentOfParentRow',
    'PercentOfParentCol',
    'PercentOfParent',
    'PercentOfRunningTotal',
  ] as const;

  it.each(percentKinds)('formats a %s value field as 0.00%% by default', async (kind) => {
    const wb = await setup();
    try {
      const result = create(wb, {
        valueFieldSettings: [{ fieldName: 'Amt', showValuesAs: PivotShowValuesAs[kind] }],
      });
      expect(result.ok).toBe(true);
      expect(dataFields(wb)).toMatchObject([
        { source: 'Amt', numberFormat: '0.00%', showValuesAs: PivotShowValuesAs[kind] },
      ]);
    } finally {
      wb.dispose();
    }
  });

  const plainKinds = [
    'Normal',
    'RunningTotalInRow',
    'RunningTotalInCol',
    'Index',
    'DifferenceFrom',
    'RankAscending',
    'RankDescending',
  ] as const;

  it.each(plainKinds)('leaves a %s value field without a number format', async (kind) => {
    const wb = await setup();
    try {
      const result = create(wb, {
        valueFieldSettings: [{ fieldName: 'Amt', showValuesAs: PivotShowValuesAs[kind] }],
      });
      expect(result.ok).toBe(true);
      expect(dataFields(wb)[0]?.numberFormat).toBe('');
    } finally {
      wb.dispose();
    }
  });

  it('lets the per-field number format override the percent default', async () => {
    const wb = await setup();
    try {
      create(wb, {
        valueNumberFormat: '#,##0',
        valueFieldSettings: [
          {
            fieldName: 'Amt',
            showValuesAs: PivotShowValuesAs.PercentOfTotal,
            numberFormat: '0.0%',
          },
        ],
      });
      expect(dataFields(wb)[0]?.numberFormat).toBe('0.0%');
    } finally {
      wb.dispose();
    }
  });

  it('lets the report number format override the percent default', async () => {
    const wb = await setup();
    try {
      create(wb, {
        valueNumberFormat: '#,##0',
        valueFieldSettings: [{ fieldName: 'Amt', showValuesAs: PivotShowValuesAs.PercentOfTotal }],
      });
      expect(dataFields(wb)[0]?.numberFormat).toBe('#,##0');
    } finally {
      wb.dispose();
    }
  });

  it('treats a blank number format as absent and falls through to the percent default', async () => {
    const wb = await setup();
    try {
      create(wb, {
        valueNumberFormat: '  ',
        valueFieldSettings: [
          { fieldName: 'Amt', showValuesAs: PivotShowValuesAs.PercentOfRow, numberFormat: ' ' },
        ],
      });
      expect(dataFields(wb)[0]?.numberFormat).toBe('0.00%');
    } finally {
      wb.dispose();
    }
  });

  it('applies the percent default per field, not to the whole report', async () => {
    const wb = await setup();
    try {
      create(wb, {
        valueFields: ['Amt', 'Channel'],
        valueFieldSettings: [{ fieldName: 'Amt', showValuesAs: PivotShowValuesAs.PercentOfTotal }],
      });
      expect(dataFields(wb).map((d) => [d.source, d.numberFormat])).toEqual([
        ['Amt', '0.00%'],
        ['Channel', ''],
      ]);
    } finally {
      wb.dispose();
    }
  });
});
