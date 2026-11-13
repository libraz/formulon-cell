import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { pivotFieldFilterModel } from '../../../src/commands/pivot-filter.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotFilterKind } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { PIVOT_ALL_DATES_PERIODS } from '../../../src/interact/pivot-filter-kinds.js';
import { openPivotFieldFilter } from '../../../src/interact/pivot-filter-panel.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

// 2026-05-01, 2026-05-15 and 2026-05-31 in the 1900 date system.
const MAY_1 = 46143;
const MAY_15 = 46157;
const MAY_31 = 46173;

/** A report whose row field is `Date` (date-formatted serials) or `Region` (text). */
const setup = async (rowField: 'Date' | 'Region', platform?: 'mac') => {
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
  if (platform) host.dataset.fcPlatform = platform;
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

  it('lists the column filter date menu in place of the label list', async () => {
    await setup('Date');
    const c = en.workbookObjects.pivotFilterPanel.conditions;
    const items = [...byLabel().options].map((o) =>
      o.disabled ? '-' : o.value === '' ? 'Choose One' : o.textContent,
    );
    const p = en.datePeriods;
    expect(items).toEqual([
      'Choose One',
      '-',
      c.equals,
      '-',
      c.before,
      c.after,
      c.between,
      '-',
      p.tomorrow,
      p.today,
      p.yesterday,
      '-',
      p.nextWeek,
      p.thisWeek,
      p.lastWeek,
      '-',
      p.nextMonth,
      p.thisMonth,
      p.lastMonth,
      '-',
      p.nextQuarter,
      p.thisQuarter,
      p.lastQuarter,
      '-',
      p.nextYear,
      p.thisYear,
      p.lastYear,
      '-',
      p.yearToDate,
      '-',
      c.allDatesInPeriod,
    ]);
    expect(
      [...byLabel().options].some((o) => o.value === String(PivotFilterKind.LabelContains)),
    ).toBe(false);
  });

  it('picks All Dates in the Period from a quarter and month pop-up in the dialog', async () => {
    const { wb } = await setup('Date');
    choose(byLabel(), 'allDatesInPeriod');
    const period = dialog()?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-period="true"]',
    );
    if (!period) throw new Error('missing period pop-up');
    expect([...period.options].map((o) => o.textContent)).toEqual(
      PIVOT_ALL_DATES_PERIODS.map(([, key]) => en.datePeriods[key]),
    );
    expect(period.value).toBe(String(PivotFilterKind.Q1));
    choose(period, String(PivotFilterKind.M5));
    dialog()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await settle();
    expect(pivotFieldFilterModel(wb, 0, 0, 0)?.condition).toMatchObject({
      fieldName: 'Date',
      kind: PivotFilterKind.M5,
    });
    choose(byLabel(), 'allDatesInPeriod');
    // Reopened on the stored month.
    expect(
      dialog()?.querySelector<HTMLSelectElement>('select[data-pivot-filter-period="true"]')?.value,
    ).toBe(String(PivotFilterKind.M5));
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

describe('PivotTable filter panel dialog titles', () => {
  it('titles the dialog with the field on the default chrome and without it on the Mac', async () => {
    await setup('Date');
    choose(byLabel(), String(PivotFilterKind.DateBetween));
    expect(dialog()?.getAttribute('aria-label')).toBe('Date Filter (Date)');
    document.body.replaceChildren();
    await setup('Date', 'mac');
    choose(byLabel(), String(PivotFilterKind.DateBetween));
    expect(dialog()?.getAttribute('aria-label')).toBe(
      en.pivotTableDialog.filterDialogs.date.macTitle,
    );
  });
});
