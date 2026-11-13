import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { pivotFieldFilterModel } from '../../../src/commands/pivot-filter.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotFilterKind } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { openPivotFieldFilter } from '../../../src/interact/pivot-filter-panel.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

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
  const ctx = {
    host,
    store,
    origin: 'contextMenu' as const,
    commandId: 'pivotFilterValue',
    wb,
    history: new History(),
    strings: en,
    onChanged: () => mutators.replaceCells(store, wb.cells(0)),
  };
  openPivotFieldFilter(ctx, { sheet: 0, row: 2, col: 0 });
  return { wb };
};

const by = (which: 'label' | 'value') =>
  document.querySelector<HTMLSelectElement>(
    `select[data-pivot-filter-by="${which}"]`,
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

describe('PivotTable filter panel conditions', () => {
  it('lists every entry enabled, Not Between on the label list only', async () => {
    await setup();
    for (const which of ['label', 'value'] as const) {
      const entries = [...by(which).options].filter((o) => !/^─+$/.test(o.text));
      expect(
        entries.every((o) => !o.disabled),
        which,
      ).toBe(true);
      expect(entries.some((o) => o.value.startsWith('unsupported'))).toBe(false);
    }
    const labels = [...by('label').options].map((o) => o.text);
    expect(labels[labels.indexOf('Between') + 1]).toBe('Not Between');
    const values = [...by('value').options].map((o) => (/^─+$/.test(o.text) ? '-' : o.text));
    expect(values).toEqual([
      'Choose One',
      '-',
      'Equals',
      'Does Not Equal',
      '-',
      'Greater Than',
      'Greater Than or Equal To',
      'Less Than',
      'Less Than or Equal To',
      '-',
      'Between',
      '-',
      'Top 10',
      'Bottom 10',
    ]);
  });

  it('opens the Top 10 dialog with Bottom preselected for Bottom 10', async () => {
    await setup();
    choose(by('value'), String(PivotFilterKind.ValueBottomCount));
    const d = dialog();
    expect(d?.textContent).toContain('Top 10 Filter (Region)');
    expect(
      d?.querySelector<HTMLSelectElement>('select[data-pivot-filter-top10-direction="true"]')
        ?.value,
    ).toBe('bottom');
  });

  it('commits a filter of the chosen label kind on OK', async () => {
    const { wb } = await setup();
    choose(by('label'), String(PivotFilterKind.LabelNotBeginsWith));
    const input = dialog()?.querySelector<HTMLInputElement>(
      'input[data-pivot-filter-value="true"]',
    );
    if (!input) throw new Error('missing label input');
    input.value = 'E';
    dialog()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await settle();
    expect(pivotFieldFilterModel(wb, 0, 0, 0)?.condition?.kind).toBe(
      PivotFilterKind.LabelNotBeginsWith,
    );
  });

  it('commits Not Between through the Value Filter dialog opened from Between', async () => {
    const { wb } = await setup();
    choose(by('value'), String(PivotFilterKind.ValueBetween));
    const d = dialog();
    const operator = d?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-operator="true"]',
    );
    const low = d?.querySelector<HTMLInputElement>('input[data-pivot-filter-value="true"]');
    const high = d?.querySelector<HTMLInputElement>('input[data-pivot-filter-value-high="true"]');
    if (!operator || !low || !high) throw new Error('missing range controls');
    choose(operator, String(PivotFilterKind.ValueNotBetween));
    low.value = '400';
    high.value = '600';
    d?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await settle();
    expect(pivotFieldFilterModel(wb, 0, 0, 0)?.condition?.kind).toBe(
      PivotFilterKind.ValueNotBetween,
    );
  });
});
