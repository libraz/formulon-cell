import { describe, expect, it } from 'vitest';
import { en } from '../../../../src/i18n/strings.js';
import {
  buildPivotEntries,
  type MenuEntry,
  type PivotMenuTarget,
} from '../../../../src/interact/context-menu-spec.js';

const target = (dateField: boolean): PivotMenuTarget => ({
  cell: 'item',
  fieldName: 'Date',
  itemName: '2026-05-15',
  dataFieldName: 'Sum of Amt',
  itemOnColumns: false,
  subtotalOn: false,
  filtered: false,
  itemPosition: 0,
  itemCount: 3,
  fieldPosition: 0,
  fieldCount: 1,
  grouped: false,
  dateField,
  canRefresh: true,
  canOpenFieldList: true,
});

const labelOf = (entries: readonly MenuEntry[], id: string): string | undefined => {
  for (const entry of entries) {
    if (entry.kind === 'item' && entry.id === id) return entry.label;
    if (entry.kind === 'submenu') {
      const found = labelOf(entry.children, id);
      if (found !== undefined) return found;
    }
  }
  return undefined;
};

describe('PivotTable context menu label filters', () => {
  it('reads Date Filters for a date field and Label Filters otherwise', () => {
    expect(labelOf(buildPivotEntries(en, target(true)), 'pivotFilterLabel')).toBe(
      en.contextMenu.pivotFilterDate,
    );
    expect(labelOf(buildPivotEntries(en, target(false)), 'pivotFilterLabel')).toBe(
      en.contextMenu.pivotFilterLabel,
    );
  });
});
