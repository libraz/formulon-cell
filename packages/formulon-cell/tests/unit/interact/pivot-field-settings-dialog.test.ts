import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotAxis, PivotShowValuesAs } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Amt', 'Qty'],
    ['East', 'Desk', 12, 1],
    ['West', 'Chair', 5, 3],
    ['East', 'Chair', 7, 2],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 },
    destination: { sheet: 0, row: 8, col: 0 },
    rowField: 'Region',
    valueFields: ['Amt'],
    valueField: 'Amt',
  });
  expect(created.ok).toBe(true);
  return wb;
};

const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('PivotTable Field dialog', () => {
  let host: HTMLElement;
  beforeEach(() => {
    document.body.replaceChildren();
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => {
    document.body.replaceChildren();
  });

  const openPane = (wb: WorkbookHandle, strings = en, history: History | null = null) => {
    const handle = attachWorkbookObjectsPanel({ host, wb, strings, history });
    expect(handle.openPivotFieldList(0, 0)).toBe(true);
    return handle;
  };
  const chipInfo = (area: string, field: string): HTMLButtonElement => {
    const button = host.querySelector<HTMLButtonElement>(
      `[data-pivot-area="${area}"] [data-pivot-area-field="${field}"] .fc-pivotdlg__area-settings`,
    );
    if (!button) throw new Error(`missing ${field} settings button`);
    return button;
  };
  const dialog = (): HTMLElement => {
    const el = document.querySelector<HTMLElement>('.fc-pvfsdlg:not([hidden])');
    if (!el) throw new Error('dialog not open');
    return el;
  };
  const options = (selector: string): string[] =>
    Array.from(dialog().querySelectorAll<HTMLElement>(`${selector} [role="option"]`)).map(
      (option) => option.textContent ?? '',
    );
  const click = (selector: string): void => {
    const el = dialog().querySelector<HTMLElement>(selector);
    if (!el) throw new Error(`missing ${selector}`);
    el.click();
  };

  it('lays out the value field dialog in the desktop order and wording', async () => {
    const wb = await setup();
    const handle = openPane(wb, ja);
    chipInfo('values', 'Amt').click();
    const root = dialog();
    expect(root.getAttribute('aria-label')).toBe('ピボットテーブル フィールド');
    expect(root.querySelector('.fc-fmtdlg__header')?.textContent).toBe(
      'ピボットテーブル フィールド',
    );
    const labels = Array.from(root.querySelectorAll('.fc-pvfsdlg__name-row > label')).map(
      (label) => label.textContent,
    );
    expect(labels).toEqual(['フィールド名:', 'フィールド名:']);
    expect(root.querySelector('[data-pvfs-source]')?.textContent).toBe('Amt');
    const name = root.querySelector<HTMLInputElement>('[data-pvfs-name]');
    expect(name?.value).toBe(wb.getPivotDefinition(0, 0)?.dataFields[0]?.name);
    expect(
      Array.from(root.querySelectorAll('[role="tab"]')).map((tab) => [
        tab.textContent,
        tab.getAttribute('aria-selected'),
      ]),
    ).toEqual([
      ['集計の方法', 'true'],
      ['計算の種類', 'false'],
    ]);
    expect(options('[data-pvfs-aggregation]')).toEqual([
      '合計',
      '個数',
      '平均',
      '最大',
      '最小',
      '積',
      '数値の個数',
      '標本標準偏差',
      '標準偏差',
      '標本分散',
      '分散',
    ]);
    expect(root.querySelector('[data-pvfs-aggregation] [aria-selected="true"]')?.textContent).toBe(
      '合計',
    );
    const showAs = root.querySelector<HTMLSelectElement>('select[data-pvfs-show-as]');
    expect(Array.from(showAs?.options ?? []).map((option) => option.textContent)).toEqual([
      '計算なし',
      '基準値との差分',
      '基準値に対する比率',
      '基準値との差分の比率',
      '累計',
      '行集計に対する比率',
      '列集計に対する比率',
      '総計に対する比率',
      '指数 (インデックス)',
      '親行集計に対する比率',
      '親列集計に対する比率',
      '親集計に対する比率',
      '比率の累計',
      '昇順での順位',
      '降順での順位',
    ]);
    expect(options('[data-pvfs-base-field]')).toEqual(['Region', 'Product', 'Amt', 'Qty']);
    expect(root.querySelector('[data-pvfs-base-field]')?.getAttribute('aria-disabled')).toBe(
      'true',
    );
    expect(options('[data-pvfs-base-item]')).toEqual([]);
    const footer = Array.from(root.querySelectorAll('.fc-pvfsdlg__footer button')).map(
      (button) => button.textContent,
    );
    expect(footer).toEqual(['表示形式...', 'キャンセル', 'OK']);
    click('[data-pvfs-cancel]');
    handle.detach();
    wb.dispose();
  });

  it('applies a new function at once, renames the header after it, and undoes in one step', async () => {
    const wb = await setup();
    const history = new History();
    const handle = openPane(wb, en, history);
    const before = wb.getPivotDefinition(0, 0)?.dataFields[0];
    chipInfo('values', 'Amt').click();
    click('[data-pvfs-aggregation] [data-pvfs-value="2"]');
    click('[data-pvfs-ok]');
    await flush();
    const after = wb.getPivotDefinition(0, 0)?.dataFields[0];
    expect(after?.aggregation).toBe(PivotAggregation.Average);
    expect(after?.name).not.toBe(before?.name);
    expect(after?.name).toMatch(/Average/);
    expect(
      host.querySelector('[data-pivot-area="values"] [data-pivot-area-field="Amt"] strong')
        ?.textContent,
    ).toBe(after?.name);
    expect(history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]).toMatchObject({
      aggregation: PivotAggregation.Sum,
      name: before?.name,
    });
    handle.detach();
    wb.dispose();
  });

  it('renames the name field live as the function changes, keeping a typed name', async () => {
    const wb = await setup();
    const handle = openPane(wb);
    const before = JSON.stringify(wb.getPivotDefinition(0, 0));
    chipInfo('values', 'Amt').click();
    expect(JSON.stringify(wb.getPivotDefinition(0, 0)), 'opening writes nothing').toBe(before);
    const name = dialog().querySelector<HTMLInputElement>('[data-pvfs-name]');
    if (!name) throw new Error('missing name');
    const sumName = name.value;
    click(`[data-pvfs-aggregation] [data-pvfs-value="${PivotAggregation.Average}"]`);
    expect(name.value).not.toBe(sumName);
    expect(name.value).toMatch(/Average/);
    const averageName = name.value;
    click(`[data-pvfs-aggregation] [data-pvfs-value="${PivotAggregation.Sum}"]`);
    expect(name.value).toBe(sumName);
    name.value = 'Mine';
    name.dispatchEvent(new Event('input', { bubbles: true }));
    click(`[data-pvfs-aggregation] [data-pvfs-value="${PivotAggregation.Max}"]`);
    expect(name.value).toBe('Mine');
    click('[data-pvfs-cancel]');
    expect(averageName).toBeTruthy();
    handle.detach();
    wb.dispose();
  });

  it('opens number-only Format Cells from Number Format and writes the code', async () => {
    const wb = await setup();
    const { createSpreadsheetStore } = await import('../../../src/store/store.js');
    const handle = attachWorkbookObjectsPanel({
      host,
      wb,
      strings: ja,
      store: createSpreadsheetStore(),
    });
    expect(handle.openPivotFieldList(0, 0)).toBe(true);
    chipInfo('values', 'Amt').click();
    const formatBtn = dialog().querySelector<HTMLButtonElement>('[data-pvfs-number-format]');
    expect(formatBtn?.disabled).toBe(false);
    formatBtn?.click();
    const fmt = Array.from(document.querySelectorAll<HTMLElement>('.fc-fmtdlg')).find(
      (el) => !el.hidden && el.querySelector('button[data-fc-tab]'),
    );
    expect(fmt?.querySelector('.fc-fmtdlg__title, .fc-fmtdlg__header')?.textContent).toBe(
      'セルの書式設定',
    );
    expect(
      Array.from(fmt?.querySelectorAll<HTMLButtonElement>('button[data-fc-tab]') ?? [])
        .filter((tab) => !tab.hidden)
        .map((tab) => tab.textContent),
    ).toEqual(['表示形式']);
    fmt?.querySelector<HTMLButtonElement>('button[data-fc-cat="percent"]')?.click();
    fmt?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    click('[data-pvfs-ok]');
    await flush();
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]?.numberFormat).toMatch(/%/);
    handle.detach();
    wb.dispose();
  });

  it('keeps a typed name and writes the calculation and its base field', async () => {
    const wb = await setup();
    const handle = openPane(wb);
    chipInfo('values', 'Amt').click();
    const name = dialog().querySelector<HTMLInputElement>('[data-pvfs-name]');
    if (!name) throw new Error('missing name');
    name.value = 'Amount share';
    name.dispatchEvent(new Event('input', { bubbles: true }));
    click('[data-pvfs-tab="showAs"]');
    expect(dialog().dataset.pvfsTab).toBe('showAs');
    const showAs = dialog().querySelector<HTMLSelectElement>('select[data-pvfs-show-as]');
    if (!showAs) throw new Error('missing show-as');
    showAs.value = String(PivotShowValuesAs.PercentOfTotal);
    showAs.dispatchEvent(new Event('change', { bubbles: true }));
    click('[data-pvfs-ok]');
    await flush();
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]).toMatchObject({
      name: 'Amount share',
      showValuesAs: PivotShowValuesAs.PercentOfTotal,
    });
    handle.detach();
    wb.dispose();
  });

  it('enables base field and item for Difference From', async () => {
    const wb = await setup();
    const handle = openPane(wb);
    chipInfo('values', 'Amt').click();
    const showAs = dialog().querySelector<HTMLSelectElement>('select[data-pvfs-show-as]');
    if (!showAs) throw new Error('missing show-as');
    showAs.value = String(PivotShowValuesAs.DifferenceFrom);
    showAs.dispatchEvent(new Event('change', { bubbles: true }));
    expect(dialog().querySelector('[data-pvfs-base-field]')?.getAttribute('aria-disabled')).toBe(
      'false',
    );
    expect(options('[data-pvfs-base-item]').slice(0, 2)).toEqual(['(previous)', '(next)']);
    click('[data-pvfs-ok]');
    await flush();
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]).toMatchObject({
      showValuesAs: PivotShowValuesAs.DifferenceFrom,
      showAsBaseField: 0,
    });
    handle.detach();
    wb.dispose();
  });

  it('shows the row field dialog with subtotals, and writes custom functions', async () => {
    const wb = await setup();
    const handle = openPane(wb, ja);
    chipInfo('rows', 'Region').click();
    const root = dialog();
    expect(root.dataset.pvfsKind).toBe('field');
    expect(root.querySelector('[data-pvfs-source]')).toBeNull();
    expect(root.querySelector<HTMLInputElement>('[data-pvfs-name]')?.value).toBe('Region');
    expect(root.querySelector('.fc-pvfsdlg__heading')?.textContent).toBe('小計');
    expect(
      Array.from(root.querySelectorAll('.fc-pvfsdlg__radio')).map((label) => label.textContent),
    ).toEqual(['自動', '指定', 'なし']);
    expect(root.querySelector<HTMLInputElement>('[data-pvfs-subtotal="auto"]')?.checked).toBe(true);
    expect(root.querySelector<HTMLInputElement>('[data-pvfs-subtotal="none"]')?.disabled).toBe(
      true,
    );
    expect(root.querySelector<HTMLInputElement>('[data-pvfs-show-no-data]')?.disabled).toBe(true);
    expect(root.querySelector('.fc-pvfsdlg__check')?.textContent).toBe(
      'データのないアイテムを表示する',
    );
    expect(options('[data-pvfs-subtotal-fns]')).toEqual([
      '合計',
      '個数',
      '平均',
      '最大',
      '最小',
      '積',
      '数値の個数',
      '標本標準偏差',
      '標準偏差',
      '標本分散',
      '分散',
    ]);
    expect(
      Array.from(root.querySelectorAll('.fc-pvfsdlg__footer button')).map((b) => b.textContent),
    ).toEqual(['非表示', 'キャンセル', 'OK']);
    click('[data-pvfs-subtotal-fns] [data-pvfs-value="0"]');
    expect(root.querySelector<HTMLInputElement>('[data-pvfs-subtotal="custom"]')?.checked).toBe(
      true,
    );
    click('[data-pvfs-ok]');
    await flush();
    expect(wb.getPivotField(0, 0, 0)).toMatchObject({ subtotalFns: [PivotAggregation.Sum] });
    expect(wb.getPivotDefinition(0, 0)?.rowFieldOrder).toEqual([0]);
    handle.detach();
    wb.dispose();
  });

  it('takes the field out of the report from Hide', async () => {
    const wb = await setup();
    const handle = openPane(wb);
    chipInfo('rows', 'Region').click();
    click('[data-pvfs-hide]');
    await flush();
    expect(wb.getPivotDefinition(0, 0)?.fields[0]?.axis).toBeNull();
    expect(wb.getPivotDefinition(0, 0)?.fields[2]?.axis).toBe(PivotAxis.Value);
    handle.detach();
    wb.dispose();
  });

  it('cancels without writing anything', async () => {
    const wb = await setup();
    const handle = openPane(wb);
    const before = JSON.stringify(wb.getPivotDefinition(0, 0));
    chipInfo('values', 'Amt').click();
    click('[data-pvfs-aggregation] [data-pvfs-value="3"]');
    click('[data-pvfs-cancel]');
    await flush();
    expect(JSON.stringify(wb.getPivotDefinition(0, 0))).toBe(before);
    expect(document.querySelector('.fc-pvfsdlg')).toBeNull();
    handle.detach();
    wb.dispose();
  });
});
