import { ComparisonOperator, DynamicFilterType, FilterKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { hydrateAutoFilterFromEngine } from '../../../src/engine/auto-filter-sync.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const VALUES: (string | number | null)[] = ['Num', 5, 50, 500, 'a1', 'b2', 'apple', null, 0];
const RANGE = { sheet: 0, r0: 0, c0: 0, r1: VALUES.length - 1, c1: 0 };

/** A book whose column holds the custom criterion '>a*', saved and opened again. */
const reopenedWithWildcardOrdering = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  try {
    VALUES.forEach((value, row) => {
      if (typeof value === 'number') wb.setNumber({ sheet: 0, row, col: 0 }, value);
      else if (typeof value === 'string') wb.setText({ sheet: 0, row, col: 0 }, value);
    });
    wb.setAutoFilter(0, {
      range: { firstRow: 0, firstCol: 0, lastRow: RANGE.r1, lastCol: 0 },
      columns: [
        {
          columnOffset: 0,
          hiddenButton: false,
          showButton: true,
          kind: FilterKind.Custom,
          filterBlank: false,
          values: [],
          dateGroups: [],
          customAnd: false,
          conditions: [{ operator: ComparisonOperator.GreaterThan, value: 'a*' }],
          top: false,
          percent: false,
          topValue: 0,
          filterValue: null,
          dynamicType: DynamicFilterType.Null,
          dynamicValue: null,
          dynamicMaxValue: null,
          valueIso: null,
          maxValueIso: null,
          dxfIndex: null,
          cellColor: false,
          iconSet: 0,
          iconIndex: null,
        },
      ],
      sort: null,
    });
    wb.applyAutoFilter(0);
    return await WorkbookHandle.loadBytes(wb.save());
  } finally {
    wb.dispose();
  }
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('Mac column filter over a criterion it cannot show', () => {
  it('opens a wildcard ordering criterion as the desktop does: empty rows, And, nothing checked', async () => {
    const wb = await reopenedWithWildcardOrdering();
    try {
      const store = createSpreadsheetStore();
      mutators.replaceCells(store, wb.cells(0));
      hydrateLayoutFromEngine(wb, store, 0);
      hydrateAutoFilterFromEngine(wb, store, 0);
      // The criterion stays the engine's: only b2 shows.
      expect([...store.getState().layout.hiddenRows].sort((a, b) => a - b)).toEqual([
        1, 2, 3, 4, 6, 7, 8,
      ]);

      const host = document.createElement('div');
      host.className = 'fc-host';
      host.dataset.fcPlatform = 'mac';
      document.body.appendChild(host);
      attachFilterDropdown({ store, host, strings: ja, locale: 'ja-JP', getWb: () => wb }).open(
        RANGE,
        0,
        { x: 10, y: 20, h: 24 },
      );
      const panel = document.querySelector<HTMLElement>('.fc-filter-dropdown') as HTMLElement;
      const select = (name: string) =>
        panel.querySelector<HTMLSelectElement>(`.fc-filter-dropdown__${name}`) as HTMLSelectElement;
      const field = (name: string) =>
        panel.querySelector<HTMLInputElement>(`.fc-filter-dropdown__${name}`) as HTMLInputElement;
      const m = ja.mac.filterPanel;
      const labels = (s: HTMLSelectElement) =>
        [...s.options].map((o) => (o.disabled ? '-' : o.textContent));

      expect(select('condition-op').value).toBe('');
      expect(field('condition-value').value).toBe('');
      expect(select('condition-op2').hidden).toBe(false);
      expect(select('condition-op2').value).toBe('');
      const join = panel.querySelector<HTMLElement>('.fc-filter-dropdown__condition-join');
      expect(join?.hidden).toBe(false);
      expect(join?.querySelector<HTMLInputElement>('input[value="and"]')?.checked).toBe(true);
      // A column of mostly numbers takes the number lists; the second row stops at Less Than Or Equal To.
      expect(labels(select('condition-op'))).toEqual([
        m.chooseOne,
        '-',
        m.equals,
        m.notEquals,
        '-',
        m.greaterThan,
        m.greaterThanOrEqual,
        m.lessThan,
        m.lessThanOrEqual,
        m.between,
        '-',
        m.top10,
        m.bottom10,
        m.aboveAverage,
        m.belowAverage,
      ]);
      expect(labels(select('condition-op2')).at(-1)).toBe(m.lessThanOrEqual);

      const boxes = [
        ...panel.querySelectorAll<HTMLInputElement>(
          '.fc-filter-dropdown__list input[type="checkbox"]',
        ),
      ];
      expect(boxes.map((box) => box.closest('label')?.textContent?.trim())).toEqual([
        m.selectAll,
        '0',
        '5',
        '50',
        '500',
        'a1',
        'apple',
        'b2',
        m.blanks,
      ]);
      expect(boxes.map((box) => box.checked)).toEqual(boxes.map(() => false));
    } finally {
      wb.dispose();
    }
  });
});
