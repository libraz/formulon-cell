import { beforeEach, describe, expect, it } from 'vitest';
import type { PivotSourceField } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotShowValuesAs } from '../../../src/engine/types.js';
import { createPivotFieldAreas } from '../../../src/interact/pivot-table-field-areas.js';
import { createDialogSelect } from '../../../src/toolbar/dialogs/form-controls.js';

const field = (name: string, index: number, numericCount = 0): PivotSourceField => ({
  name,
  index,
  numericCount,
});

const FIELDS = [
  field('Region', 0),
  field('Product', 1),
  field('Channel', 2),
  field('Zone', 3),
  field('Sales', 4, 5),
  field('Qty', 5, 5),
];

const setup = () => {
  const make = () => createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const selects = {
    rowSelect: make(),
    colSelect: make(),
    filterSelect: make(),
    valueSelect: make(),
  };
  const areas = createPivotFieldAreas(selects);
  areas.resetSource(FIELDS, 'None');
  return { areas, ...selects };
};

describe('createPivotFieldAreas resetSource', () => {
  it('seeds row, empty column and first numeric value on a fresh source', () => {
    const { areas, rowSelect, colSelect, filterSelect, valueSelect } = setup();
    expect(rowSelect.value).toBe('Region');
    expect(colSelect.value).toBe('');
    expect(filterSelect.value).toBe('');
    expect(valueSelect.value).toBe('Sales');
    expect(areas.valueFields).toEqual(['Sales']);
    expect(areas.filterFields).toEqual([]);
    expect(Array.from(colSelect.options).map((o) => o.value)).toEqual([
      '',
      ...FIELDS.map((f) => f.name),
    ]);
    expect(Array.from(valueSelect.options).map((o) => o.value)).toEqual(['Sales', 'Qty']);
  });

  it('falls back to every field as a value candidate when none is numeric', () => {
    const { areas, valueSelect } = setup();
    areas.resetSource([field('A', 0), field('B', 1)], 'None');
    expect(Array.from(valueSelect.options).map((o) => o.value)).toEqual(['A', 'B']);
    expect(areas.valueFields).toEqual(['B']);
  });

  it('keeps surviving assignments and drops settings of fields that disappeared', () => {
    const { areas, rowSelect, valueSelect } = setup();
    areas.assignFieldToArea('Channel', 'filters', FIELDS);
    areas.assignFieldToArea('Zone', 'filters', FIELDS);
    areas.assignFieldToArea('Qty', 'values', FIELDS);
    areas.setFilterCondition('Channel', { kind: 'label-contains', value: 'web' });
    areas.setFilterCondition('Zone', { kind: 'label-contains', value: 'north' });
    areas.setFilterItemVisibility('Zone', 'N', false);
    areas.setValueFieldSetting('Qty', { numberFormat: '0.0' });
    areas.setValueFieldSetting('Sales', { aggregation: PivotAggregation.Max });
    expect(areas.valueFields).toEqual(['Sales', 'Qty']);

    areas.resetSource(
      FIELDS.filter((f) => f.name !== 'Zone' && f.name !== 'Qty'),
      'None',
    );

    expect(rowSelect.value).toBe('Region');
    expect(areas.filterFields).toEqual(['Channel']);
    expect(areas.valueFields).toEqual(['Sales']);
    expect(valueSelect.value).toBe('Sales');
    expect(areas.selectedFilterCondition('Channel')).toEqual({
      kind: 'label-contains',
      value: 'web',
    });
    expect(areas.selectedFilterCondition('Zone')).toBeUndefined();
    expect(areas.selectedValueSetting('Qty')).toEqual({});
    expect(areas.selectedValueSetting('Sales')).toEqual({ aggregation: PivotAggregation.Max });
    expect(areas.filterItemVisibility('Zone', 'N')).toBe(true);
  });

  it('replaces a vanished row field with the first field', () => {
    const { areas, rowSelect } = setup();
    areas.assignFieldToArea('Product', 'rows', FIELDS);
    expect(rowSelect.value).toBe('Product');
    areas.resetSource([field('Alpha', 0), field('Beta', 1, 3)], 'None');
    expect(rowSelect.value).toBe('Alpha');
  });
});

describe('createPivotFieldAreas toggleFromFieldList', () => {
  let ctx: ReturnType<typeof setup>;
  beforeEach(() => {
    ctx = setup();
  });

  it('checks a numeric field into values', () => {
    ctx.areas.toggleFromFieldList('Qty', true, FIELDS);
    expect(ctx.areas.valueFields).toEqual(['Sales', 'Qty']);
  });

  it('checks a text field into the row when the row is empty', () => {
    ctx.rowSelect.value = '';
    ctx.areas.toggleFromFieldList('Product', true, FIELDS);
    expect(ctx.rowSelect.value).toBe('Product');
  });

  it('checks a text field into the column when row is taken and column is empty', () => {
    ctx.areas.toggleFromFieldList('Product', true, FIELDS);
    expect(ctx.colSelect.value).toBe('Product');
  });

  it('then fills filters once row and column are taken', () => {
    ctx.areas.toggleFromFieldList('Product', true, FIELDS);
    ctx.areas.toggleFromFieldList('Channel', true, FIELDS);
    ctx.areas.toggleFromFieldList('Zone', true, FIELDS);
    expect(ctx.areas.filterFields).toEqual(['Channel', 'Zone']);
    expect(ctx.filterSelect.value).toBe('Channel');
  });

  it('unchecks a filter field and drops its condition', () => {
    ctx.areas.assignFieldToArea('Channel', 'filters', FIELDS);
    ctx.areas.setFilterCondition('Channel', { kind: 'label-contains', value: 'web' });
    ctx.areas.toggleFromFieldList('Channel', false, FIELDS);
    expect(ctx.areas.filterFields).toEqual([]);
    expect(ctx.filterSelect.value).toBe('');
    expect(ctx.areas.selectedFilterCondition('Channel')).toBeUndefined();
  });

  it('unchecks the column field to empty', () => {
    ctx.areas.assignFieldToArea('Product', 'columns', FIELDS);
    ctx.areas.toggleFromFieldList('Product', false, FIELDS);
    expect(ctx.colSelect.value).toBe('');
  });

  it('unchecks the row field by moving the row to the first other field', () => {
    ctx.areas.toggleFromFieldList('Region', false, FIELDS);
    expect(ctx.rowSelect.value).toBe('Product');
  });

  it('unchecks the last value by promoting another candidate', () => {
    ctx.areas.toggleFromFieldList('Sales', false, FIELDS);
    expect(ctx.areas.valueFields).toEqual(['Qty']);
    expect(ctx.valueSelect.value).toBe('Qty');
  });

  it('unchecks one of several values without touching the rest', () => {
    ctx.areas.toggleFromFieldList('Qty', true, FIELDS);
    ctx.areas.setValueFieldSetting('Qty', { numberFormat: '0.0' });
    ctx.areas.toggleFromFieldList('Qty', false, FIELDS);
    expect(ctx.areas.valueFields).toEqual(['Sales']);
    expect(ctx.areas.selectedValueSetting('Qty')).toEqual({});
  });
});

describe('createPivotFieldAreas assignment helpers', () => {
  it('moves a field between areas and ignores unknown fields', () => {
    const { areas, rowSelect, colSelect } = setup();
    areas.assignFieldToArea('Region', 'columns', FIELDS);
    expect(colSelect.value).toBe('Region');
    expect(rowSelect.value).toBe('');
    areas.assignFieldToArea('Missing', 'rows', FIELDS);
    expect(rowSelect.value).toBe('');
  });

  it('refuses to drop a non-numeric field into values', () => {
    const { areas } = setup();
    areas.assignFieldToArea('Product', 'values', FIELDS);
    expect(areas.valueFields).toEqual(['Sales']);
  });

  it('prunes the active settings target once its field leaves the area', () => {
    const { areas } = setup();
    areas.setActive({ kind: 'rows', fieldName: 'Region' });
    areas.pruneActiveSettings();
    expect(areas.active).toEqual({ kind: 'rows', fieldName: 'Region' });
    areas.assignFieldToArea('Region', 'columns', FIELDS);
    areas.pruneActiveSettings();
    expect(areas.active).toBeNull();
  });

  it('collapses to a single value or filter from the select', () => {
    const { areas, valueSelect, filterSelect } = setup();
    areas.toggleFromFieldList('Qty', true, FIELDS);
    valueSelect.value = 'Qty';
    areas.selectSingleValue();
    expect(areas.valueFields).toEqual(['Qty']);
    filterSelect.value = 'Zone';
    areas.selectSingleFilter();
    expect(areas.filterFields).toEqual(['Zone']);
    filterSelect.value = '';
    areas.selectSingleFilter();
    expect(areas.filterFields).toEqual([]);
  });

  it('drops a value setting that holds no information', () => {
    const { areas } = setup();
    areas.setValueFieldSetting('Sales', { numberFormat: ' 0.0 ' });
    expect(areas.selectedValueSetting('Sales')).toEqual({ numberFormat: '0.0' });
    areas.setValueFieldSetting('Sales', { numberFormat: '  ' });
    expect(areas.selectedValueSetting('Sales')).toEqual({});
  });
});

describe('createPivotFieldAreas toCreateSpec', () => {
  it('applies defaults to value fields without their own setting', () => {
    const { areas } = setup();
    areas.toggleFromFieldList('Qty', true, FIELDS);
    areas.setValueFieldSetting('Qty', {
      aggregation: PivotAggregation.Max,
      showValuesAs: PivotShowValuesAs.PercentOfTotal,
    });
    const spec = areas.toCreateSpec({ aggregation: PivotAggregation.Sum, numberFormat: '#,##0' });
    expect(spec.valueFieldSettings).toEqual([
      { fieldName: 'Sales', aggregation: PivotAggregation.Sum, numberFormat: '#,##0' },
      {
        fieldName: 'Qty',
        aggregation: PivotAggregation.Max,
        numberFormat: '#,##0',
        showValuesAs: PivotShowValuesAs.PercentOfTotal,
      },
    ]);
  });

  it('flattens item visibility and filter conditions per filter field', () => {
    const { areas } = setup();
    areas.assignFieldToArea('Channel', 'filters', FIELDS);
    areas.setFilterItemVisibility('Channel', 'web', false);
    areas.setFilterCondition('Channel', { kind: 'label-contains', value: 'we' });
    areas.setFilterCondition('Zone', { kind: 'label-contains', value: 'ignored' });
    const spec = areas.toCreateSpec({ aggregation: PivotAggregation.Sum, numberFormat: '' });
    expect(spec.filterItems).toEqual([{ fieldName: 'Channel', itemName: 'web', visible: false }]);
    expect(spec.pivotFilters).toHaveLength(1);
    expect(spec.pivotFilters[0]?.fieldName).toBe('Channel');
  });
});
