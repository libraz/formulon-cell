import { describe, expect, it } from 'vitest';
import type { PivotSourceField } from '../../../src/commands/pivot-table.js';
import {
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PivotAggregation,
  PivotShowValuesAs,
} from '../../../src/engine/types.js';
import { en } from '../../../src/i18n/strings.js';
import {
  pivotShowAsNeedsBaseField,
  pivotShowAsNeedsBaseItem,
  renderPivotFieldSettingsPanel,
  resolvePivotShowAsBase,
} from '../../../src/interact/pivot-field-settings.js';
import { createPivotFieldAreas } from '../../../src/interact/pivot-table-field-areas.js';
import { createDialogSelect } from '../../../src/toolbar/dialogs/form-controls.js';

const FIELDS: PivotSourceField[] = [
  { name: 'Region', index: 0, numericCount: 0 },
  { name: 'Product', index: 1, numericCount: 0 },
  { name: 'Sales', index: 2, numericCount: 5 },
];
const ITEMS: Record<string, string[]> = { Region: ['East', 'West', 'North'], Product: ['A', 'B'] };

const BASE_FIELD_KINDS = [
  PivotShowValuesAs.PercentOf,
  PivotShowValuesAs.PercentOfParent,
  PivotShowValuesAs.DifferenceFrom,
  PivotShowValuesAs.PercentDifferenceFrom,
  PivotShowValuesAs.PercentOfRunningTotal,
  PivotShowValuesAs.RankAscending,
  PivotShowValuesAs.RankDescending,
];
const BASE_ITEM_KINDS = [
  PivotShowValuesAs.PercentOf,
  PivotShowValuesAs.DifferenceFrom,
  PivotShowValuesAs.PercentDifferenceFrom,
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
  selects.colSelect.value = 'Product';
  const panelEl = document.createElement('div');
  const render = (): void =>
    renderPivotFieldSettingsPanel({
      panelEl,
      active: { kind: 'values', fieldName: 'Sales' },
      strings: en.pivotTableDialog,
      fields: FIELDS,
      controls: {
        ...selects,
        rowSortSelect: make(),
        colSortSelect: make(),
        aggSelect: make(),
        numberFormatInput: document.createElement('input'),
        rowSubtotalTop: document.createElement('input'),
        colSubtotalTop: document.createElement('input'),
      },
      selectedValueFields: areas.valueFields,
      selectedFilterFields: areas.filterFields,
      selectedValueSetting: areas.selectedValueSetting,
      setValueFieldSetting: areas.setValueFieldSetting,
      fieldCanBeValue: areas.fieldCanBeValue,
      replaceFilterField: areas.replaceFilterField,
      normalizeSelectedFilters: areas.normalizeSelectedFilters,
      refreshFieldList: () => {},
      filterItemVisibility: areas.filterItemVisibility,
      setFilterItemVisibility: areas.setFilterItemVisibility,
      selectedFilterCondition: areas.selectedFilterCondition,
      setFilterCondition: areas.setFilterCondition,
      inferFilterItems: (name) => ITEMS[name] ?? [],
    });
  render();
  const selectsIn = (): HTMLSelectElement[] =>
    Array.from(panelEl.querySelectorAll<HTMLSelectElement>('select'));
  const choose = (select: HTMLSelectElement, value: number | string): void => {
    select.value = String(value);
    select.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const spec = () =>
    areas.toCreateSpec({ aggregation: PivotAggregation.Sum, numberFormat: '' })
      .valueFieldSettings[0];
  return { areas, panelEl, selectsIn, choose, spec };
};

describe('pivot show-values-as base requirements', () => {
  it('flags exactly the kinds the engine reads a base field and base item for', () => {
    const all = Object.values(PivotShowValuesAs);
    expect(all.filter((k) => pivotShowAsNeedsBaseField(k)).sort()).toEqual(
      [...BASE_FIELD_KINDS].sort(),
    );
    expect(all.filter((k) => pivotShowAsNeedsBaseItem(k)).sort()).toEqual(
      [...BASE_ITEM_KINDS].sort(),
    );
  });

  it('defaults a missing or stale base to the first eligible field and (previous)', () => {
    const rank = { showValuesAs: PivotShowValuesAs.RankDescending, showAsBaseField: 'Gone' };
    expect(resolvePivotShowAsBase(rank, ['Region', 'Product'])).toEqual({
      showValuesAs: PivotShowValuesAs.RankDescending,
      showAsBaseField: 'Region',
    });
    expect(
      resolvePivotShowAsBase({ showValuesAs: PivotShowValuesAs.PercentOf }, ['Product']),
    ).toEqual({
      showValuesAs: PivotShowValuesAs.PercentOf,
      showAsBaseField: 'Product',
      showAsBaseItem: PIVOT_SHOW_AS_BASE_PREVIOUS,
    });
    expect(
      resolvePivotShowAsBase(
        { showValuesAs: PivotShowValuesAs.PercentOf, showAsBaseField: 'Region', showAsBaseItem: 7 },
        ['Region'],
        () => 3,
      ).showAsBaseItem,
    ).toBe(PIVOT_SHOW_AS_BASE_PREVIOUS);
    expect(
      resolvePivotShowAsBase(
        { showValuesAs: PivotShowValuesAs.PercentOfTotal, showAsBaseField: 'Region' },
        ['Region'],
      ),
    ).toEqual({ showValuesAs: PivotShowValuesAs.PercentOfTotal });
  });
});

describe('pivot value field settings panel show-as controls', () => {
  it('lists every show-values-as kind once', () => {
    const { selectsIn } = setup();
    const values = Array.from(selectsIn()[1]?.options ?? []).map((o) => Number(o.value));
    expect([...values].sort((a, b) => a - b)).toEqual(
      Object.values(PivotShowValuesAs).sort((a, b) => a - b),
    );
    expect(values).toHaveLength(16);
  });

  it('shows base controls exactly for the kinds that need them', () => {
    const { selectsIn, choose } = setup();
    for (const kind of Object.values(PivotShowValuesAs)) {
      choose(selectsIn()[1] as HTMLSelectElement, kind);
      const expected =
        2 + (pivotShowAsNeedsBaseField(kind) ? 1 : 0) + (pivotShowAsNeedsBaseItem(kind) ? 1 : 0);
      // aggregation + show-as, then base field, then base item
      expect(selectsIn(), `kind ${kind}`).toHaveLength(expected);
    }
  });

  it('offers row/column fields and the field items with (previous)/(next)', () => {
    const { selectsIn, choose } = setup();
    choose(selectsIn()[1] as HTMLSelectElement, PivotShowValuesAs.DifferenceFrom);
    const [, , baseField, baseItem] = selectsIn();
    expect(Array.from(baseField?.options ?? []).map((o) => o.value)).toEqual(['Region', 'Product']);
    expect(Array.from(baseItem?.options ?? []).map((o) => [o.value, o.textContent])).toEqual([
      ['0', 'East'],
      ['1', 'West'],
      ['2', 'North'],
      [String(PIVOT_SHOW_AS_BASE_PREVIOUS), '(previous)'],
      [String(PIVOT_SHOW_AS_BASE_NEXT), '(next)'],
    ]);
    expect(baseItem?.value).toBe(String(PIVOT_SHOW_AS_BASE_PREVIOUS));
  });

  it('sends the default base field and the chosen base item in the create spec', () => {
    const { selectsIn, choose, spec } = setup();
    choose(selectsIn()[1] as HTMLSelectElement, PivotShowValuesAs.PercentOf);
    expect(spec()).toMatchObject({
      showValuesAs: PivotShowValuesAs.PercentOf,
      showAsBaseField: 'Region',
      showAsBaseItem: PIVOT_SHOW_AS_BASE_PREVIOUS,
    });
    choose(selectsIn()[3] as HTMLSelectElement, 1);
    expect(spec()).toMatchObject({ showAsBaseField: 'Region', showAsBaseItem: 1 });
    choose(selectsIn()[2] as HTMLSelectElement, 'Product');
    expect(spec()).toMatchObject({
      showAsBaseField: 'Product',
      showAsBaseItem: PIVOT_SHOW_AS_BASE_PREVIOUS,
    });
  });

  it('sends only the base field for rank kinds and drops it for plain kinds', () => {
    const { selectsIn, choose, spec } = setup();
    choose(selectsIn()[1] as HTMLSelectElement, PivotShowValuesAs.RankDescending);
    const rank = spec();
    expect(rank).toMatchObject({
      showValuesAs: PivotShowValuesAs.RankDescending,
      showAsBaseField: 'Region',
    });
    expect(rank?.showAsBaseItem).toBeUndefined();
    choose(selectsIn()[1] as HTMLSelectElement, PivotShowValuesAs.PercentOfTotal);
    expect(spec()?.showAsBaseField).toBeUndefined();
  });
});
