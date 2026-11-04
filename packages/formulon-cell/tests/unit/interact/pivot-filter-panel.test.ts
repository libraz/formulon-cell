import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setPivotFieldItemsChecked } from '../../../src/commands/pivot-filter.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { openPivotFieldFilter } from '../../../src/interact/pivot-filter-panel.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { excelRibbonIconPaths } from '../../../src/toolbar/excel-ribbon-icons.js';

/** Region / Amt at K1:L4; report at A3: rows Region, Sum of Amt (A3:B7). */
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Amt'],
    ['East', 500],
    ['North', 300],
    ['South', 700],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 3, c1: 11 },
    destination: { sheet: 0, row: 2, col: 0 },
    rowField: 'Region',
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const host = document.createElement('div');
  document.body.appendChild(host);
  const history = new History();
  let changes = 0;
  const ctx = {
    host,
    store,
    wb,
    history,
    strings: ja,
    onChanged: () => {
      changes += 1;
      mutators.replaceCells(store, wb.cells(0));
    },
  };
  return { wb, store, ctx, history, changes: () => changes };
};

const panel = () => document.querySelector<HTMLElement>('.fc-pivot-filter-panel');
const box = (value: string) =>
  panel()?.querySelector<HTMLInputElement>(`input[data-pivot-filter-item="${value}"]`);
const selectAll = () =>
  panel()?.querySelector<HTMLInputElement>('input[data-pivot-filter-select-all]');
const clearButton = () =>
  panel()?.querySelector<HTMLButtonElement>('.fc-pivot-filter-panel__clear');
const grandTotal = (wb: WorkbookHandle) =>
  [...wb.pivotCells(0)].find((c) => c.kind === 6 && c.value.kind === 'number')?.value;

afterEach(() => {
  document.body.replaceChildren();
});

describe('PivotTable filter panel', () => {
  it('opens from the Row Labels button titled with the field, listing its items', async () => {
    const { ctx } = await setup();
    expect(openPivotFieldFilter(ctx, { sheet: 0, row: 2, col: 0 })).toBe(true);
    expect(panel()?.getAttribute('aria-label')).toBe('Region');
    expect(panel()?.querySelector('.fc-pivot-filter-panel__title')?.textContent).toBe('Region');
    const labels = [...(panel()?.querySelectorAll('.fc-pivot-filter-panel__item') ?? [])].map(
      (l) => l.textContent,
    );
    expect(labels).toEqual(['(すべて選択)', 'East', 'North', 'South']);
    expect(selectAll()?.checked).toBe(true);
    expect(clearButton()?.disabled).toBe(true);
    // The desktop lists sorting, then the label and value filters, then the search box.
    expect(panel()?.textContent).toContain('昇順');
    expect(panel()?.textContent).toContain('ラベル別:');
    expect(panel()?.textContent).toContain('値別:');
  });

  it('keys the sort buttons with the small A-over-Z glyphs', async () => {
    // Excel (Mac): a blue A over a dark Z (Z over A for descending) and a
    // short dark arrow, not the ribbon's full-height accent arrow.
    const { ctx } = await setup();
    openPivotFieldFilter(ctx, { sheet: 0, row: 2, col: 0 });
    for (const [order, name] of [
      ['asc', 'sortAscSmall'],
      ['desc', 'sortDescSmall'],
    ] as const) {
      const icon = panel()?.querySelector(`[data-pivot-filter-sort="${order}"] svg`);
      const drawn = [...(icon?.querySelectorAll('path') ?? [])].map((p) => p.getAttribute('d'));
      expect(drawn).toEqual(excelRibbonIconPaths(name)?.map((segment) => segment.d));
    }
  });

  it('applies an unchecked item at once as one undo step', async () => {
    const { wb, ctx, history, changes } = await setup();
    openPivotFieldFilter(ctx, { sheet: 0, row: 2, col: 0 });
    const north = box('North');
    if (!north) throw new Error('no North');
    north.checked = false;
    north.dispatchEvent(new Event('change'));
    expect(grandTotal(wb)).toEqual({ kind: 'number', value: 1200 });
    expect(changes()).toBe(1);
    expect(selectAll()?.indeterminate).toBe(true);
    expect(clearButton()?.disabled).toBe(false);
    expect(box('North')?.checked).toBe(false);
    history.undo();
    expect(grandTotal(wb)).toEqual({ kind: 'number', value: 1500 });
  });

  it('clears the filter', async () => {
    const { wb, ctx } = await setup();
    openPivotFieldFilter(ctx, { sheet: 0, row: 2, col: 0 });
    const north = box('North');
    if (!north) throw new Error('no North');
    north.checked = false;
    north.dispatchEvent(new Event('change'));
    clearButton()?.click();
    expect(grandTotal(wb)).toEqual({ kind: 'number', value: 1500 });
    expect(selectAll()?.checked).toBe(true);
    expect(clearButton()?.disabled).toBe(true);
  });

  it('offers only the conditions the engine can apply', async () => {
    const { ctx } = await setup();
    openPivotFieldFilter(ctx, { sheet: 0, row: 2, col: 0 });
    const byLabel = panel()?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-by="label"]',
    );
    const enabled = [...(byLabel?.options ?? [])].filter((o) => !o.disabled).map((o) => o.text);
    expect(enabled).toEqual(['1 つ選択してください', 'で始まる', 'を含む']);
  });

  it('is not offered on a value cell', async () => {
    const { ctx } = await setup();
    expect(openPivotFieldFilter(ctx, { sheet: 0, row: 3, col: 1 })).toBe(false);
    expect(panel()).toBeNull();
  });
});

describe('PivotTable filter panel after a write by cache index', () => {
  it('lists every item by name, the hidden one unchecked, on reopening', async () => {
    const { wb, ctx } = await setup();
    // Write the items the way undo and a loaded report do: by cache index only.
    expect(
      setPivotFieldItemsChecked(wb, 0, 0, 0, [
        { value: '', checked: true, cacheIndex: 0 },
        { value: '', checked: false, cacheIndex: 1 },
        { value: '', checked: true, cacheIndex: 2 },
      ]),
    ).toBe(true);
    expect(grandTotal(wb)).toEqual({ kind: 'number', value: 1200 });
    expect(openPivotFieldFilter(ctx, { sheet: 0, row: 2, col: 0 })).toBe(true);
    const items = [
      ...(panel()?.querySelectorAll<HTMLInputElement>('input[data-pivot-filter-item]') ?? []),
    ];
    expect(items.map((i) => [i.dataset.pivotFilterItem, i.checked])).toEqual([
      ['East', true],
      ['North', false],
      ['South', true],
    ]);
  });
});
