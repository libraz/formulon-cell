import { describe, expect, it } from 'vitest';
import type { PivotSourceField } from '../../../src/commands/pivot-table.js';
import { en } from '../../../src/i18n/strings.js';
import { renderPivotFieldSettingsPanel } from '../../../src/interact/pivot-field-settings.js';
import { createPivotFieldAreas } from '../../../src/interact/pivot-table-field-areas.js';
import { createDialogSelect } from '../../../src/toolbar/dialogs/form-controls.js';

const FIELDS: PivotSourceField[] = [
  { name: 'Region', index: 0, numeric: false },
  { name: 'Product', index: 1, numeric: false },
  { name: 'Sales', index: 2, numeric: true },
];
const ITEMS: Record<string, string[]> = { Product: ['A', 'B'] };

const render = (): HTMLDivElement => {
  const make = () => createDialogSelect([], '', { className: 'fc-fmtdlg__select' });
  const selects = {
    rowSelect: make(),
    colSelect: make(),
    filterSelect: make(),
    valueSelect: make(),
  };
  const areas = createPivotFieldAreas(selects);
  areas.resetSource(FIELDS, 'None');
  areas.assignFieldToArea('Product', 'filters', FIELDS);
  const panelEl = document.createElement('div');
  renderPivotFieldSettingsPanel({
    panelEl,
    active: { kind: 'filters', fieldName: 'Product' },
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
    inferFilterItems: (name) => ITEMS[name] ?? [],
  });
  return panelEl;
};

describe('Create PivotTable Field Settings for a Filters field', () => {
  it('offers no filter condition controls', () => {
    const panel = render();
    expect(panel.querySelector('select[data-pivot-filter-category]')).toBeNull();
    expect(panel.querySelector('select[data-pivot-filter-condition]')).toBeNull();
    expect(Array.from(panel.querySelectorAll('button'))).toEqual([]);
  });

  it('still lists the field items as a checklist', () => {
    const panel = render();
    const items = Array.from(
      panel.querySelectorAll<HTMLInputElement>(
        '.fc-pivotdlg__settings-items input[type="checkbox"]',
      ),
    );
    expect(items.map((box) => box.checked)).toEqual([true, true]);
    expect(panel.querySelector('.fc-pivotdlg__settings-items')?.textContent).toContain('A');
  });
});
