import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { openPivotFieldFilter } from '../../../src/interact/pivot-filter-panel.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** A report with Region as its filter field and Kind on rows; opens Region's panel. */
const openPageFieldPanel = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Kind', 'Amt'],
    ['East', 'a', 500],
    ['North', 'b', 300],
    ['West', 'a', 700],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 3, c1: 12 },
    destination: { sheet: 0, row: 4, col: 0 },
    rowField: 'Kind',
    filterField: 'Region',
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const label = [...wb.cells(0)].find(
    ({ addr, value }) => addr.col === 0 && value.kind === 'text' && value.value === 'Region',
  )?.addr;
  expect(label).toBeDefined();
  const host = document.createElement('div');
  document.body.appendChild(host);
  openPivotFieldFilter(
    {
      host,
      store,
      origin: 'contextMenu' as const,
      commandId: 'pivotFilterValue',
      wb,
      history: new History(),
      strings: en,
      onChanged: () => mutators.replaceCells(store, wb.cells(0)),
    },
    { sheet: 0, row: label?.row ?? 0, col: 1 },
  );
  return wb;
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('PivotTable filter panel on a report filter field', () => {
  it('offers search, the item checklist and Clear Filter only, as the desktop does', async () => {
    const wb = await openPageFieldPanel();
    try {
      const panel = document.querySelector<HTMLElement>('.fc-pivot-filter-panel') as HTMLElement;
      expect(panel).not.toBeNull();
      expect([...panel.children].map((child) => child.className)).toEqual([
        'fc-pivot-filter-panel__title',
        'fc-pivot-filter-panel__search-box',
        'fc-pivot-filter-panel__list',
        'fc-pivot-filter-panel__footer',
      ]);
      const boxes = [...panel.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')];
      expect(boxes.map((box) => box.closest('label')?.textContent?.trim())).toEqual([
        en.workbookObjects.pivotFilterPanel.selectAll,
        'East',
        'North',
        'West',
      ]);
      expect(boxes.every((box) => box.checked)).toBe(true);
    } finally {
      wb.dispose();
    }
  });
});
