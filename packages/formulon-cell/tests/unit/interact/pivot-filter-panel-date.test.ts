import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { pivotFieldFilterModel } from '../../../src/commands/pivot-filter.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotFilterKind } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { PIVOT_FILTER_PERIODS } from '../../../src/interact/pivot-filter-kinds.js';
import { openPivotFieldFilter } from '../../../src/interact/pivot-filter-panel.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

// 2026-05-01, 2026-05-15 and 2026-05-31 in the 1900 date system.
const MAY_1 = 46143;
const MAY_15 = 46157;
const MAY_31 = 46173;

/** A report whose row field is `Date` (date-formatted serials) or `Region` (text). */
const setup = async (rowField: 'Date' | 'Region') => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Date', 'Region', 'Amt'],
    [MAY_1, 'East', 10],
    [MAY_15, 'West', 20],
    [MAY_31, 'East', 30],
  ];
  const base = wb.getCellXf(0);
  if (!base) throw new Error('no default cell format');
  const dateXf = wb.addXfRecord({ ...base, numFmtId: wb.addNumFmtCode('yyyy-mm-dd') });
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
      if (r > 0 && c === 0) wb.setCellXfIndex(0, r, 10, dateXf);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 3, c1: 12 },
    destination: { sheet: 0, row: 2, col: 0 },
    rowField,
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const host = document.createElement('div');
  document.body.appendChild(host);
  openPivotFieldFilter(
    {
      host,
      store,
      origin: 'contextMenu' as const,
      commandId: 'pivotFilterLabel',
      wb,
      history: new History(),
      strings: en,
      onChanged: () => mutators.replaceCells(store, wb.cells(0)),
    },
    { sheet: 0, row: 2, col: 0 },
  );
  return { wb };
};

const byLabel = () =>
  document.querySelector<HTMLSelectElement>(
    'select[data-pivot-filter-by="label"]',
  ) as HTMLSelectElement;
const choose = (select: HTMLSelectElement, value: string): void => {
  select.value = value;
  select.dispatchEvent(new Event('change', { bubbles: true }));
};
const dialog = () => document.querySelector<HTMLElement>('.fc-pivotdlg--filter');
const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

afterEach(() => {
  document.body.replaceChildren();
});

describe('PivotTable filter panel on a date field', () => {
  it('reads a date-only field as a date field and a text field as not', async () => {
    const { wb } = await setup('Date');
    expect(pivotFieldFilterModel(wb, 0, 0, 0)?.dateField).toBe(true);
    document.body.replaceChildren();
    const text = await setup('Region');
    const region = text.wb
      .getPivotDefinition(0, 0)
      ?.fields.findIndex((field) => field.sourceName === 'Region');
    expect(pivotFieldFilterModel(text.wb, 0, 0, region ?? -1)?.dateField).toBe(false);
  });

  // The desktop's Date Filters menu in place of Label Filters. The Mac pop-up's exact contents
  // are provisional: they follow the desktop menu until measured on the Mac app.
  it('lists the date comparisons, then the relative periods, in place of the label list', async () => {
    await setup('Date');
    const enabled = [...byLabel().options]
      .filter((o) => !o.disabled && o.value !== '')
      .map((o) => [Number(o.value), o.textContent]);
    expect(enabled).toEqual([
      [PivotFilterKind.DateEqual, en.workbookObjects.pivotFilterPanel.conditions.equals],
      [PivotFilterKind.DateOlderThan, en.workbookObjects.pivotFilterPanel.conditions.before],
      [PivotFilterKind.DateNewerThan, en.workbookObjects.pivotFilterPanel.conditions.after],
      [PivotFilterKind.DateBetween, en.workbookObjects.pivotFilterPanel.conditions.between],
      ...PIVOT_FILTER_PERIODS.map(([kind, key]) => [kind, en.datePeriods[key]]),
    ]);
    expect(
      [...byLabel().options].some((o) => o.value === String(PivotFilterKind.LabelContains)),
    ).toBe(false);
  });

  it('keeps the label list for a text field', async () => {
    await setup('Region');
    const values = [...byLabel().options].map((o) => o.value);
    expect(values).toContain(String(PivotFilterKind.LabelContains));
    expect(values).not.toContain(String(PivotFilterKind.DateBetween));
  });

  it('applies a relative period at once, with no dialog', async () => {
    const { wb } = await setup('Date');
    choose(byLabel(), String(PivotFilterKind.ThisMonth));
    expect(dialog()).toBeNull();
    expect(pivotFieldFilterModel(wb, 0, 0, 0)?.condition).toMatchObject({
      fieldName: 'Date',
      kind: PivotFilterKind.ThisMonth,
    });
  });

  it('commits Between through the Date Filter dialog', async () => {
    const { wb } = await setup('Date');
    choose(byLabel(), String(PivotFilterKind.DateBetween));
    const d = dialog();
    const low = d?.querySelector<HTMLInputElement>('input[data-pivot-filter-value="true"]');
    const high = d?.querySelector<HTMLInputElement>('input[data-pivot-filter-value-high="true"]');
    if (!low || !high) throw new Error('missing date range inputs');
    low.value = '5/10/2026';
    high.value = '5/20/2026';
    d?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await settle();
    expect(pivotFieldFilterModel(wb, 0, 0, 0)?.condition).toMatchObject({
      kind: PivotFilterKind.DateBetween,
      value: 46152,
      valueHigh: 46162,
    });
  });
});
