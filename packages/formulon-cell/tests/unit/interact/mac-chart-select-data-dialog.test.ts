import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createSessionChart, sessionChartById } from '../../../src/commands/session-chart.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import {
  attachMacChartSelectDataDialog,
  DEFAULT_SELECT_DATA_LABELS,
  type MacChartSelectDataHandle,
  type SelectDataLabels,
} from '../../../src/interact/mac-chart-select-data-dialog.js';
import { DEFAULT_LABELS as SETTINGS_DEFAULT_LABELS } from '../../../src/interact/session-chart-settings-labels.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });

/** North / South over Jan-Mar in A1:C4, a West column in D1:D4. */
const seed = (wb: WorkbookHandle): void => {
  const rows: (string | number)[][] = [
    ['', 'North', 'South', 'West'],
    ['Jan', 10, 20, 30],
    ['Feb', 15, 25, 35],
    ['Mar', 12, 22, 32],
  ];
  rows.forEach((cells, row) => {
    cells.forEach((value, col) => {
      if (typeof value === 'number') wb.setNumber({ sheet: 0, row, col }, value);
      else if (value !== '') wb.setText({ sheet: 0, row, col }, value);
    });
  });
};

const $ = <T extends Element = HTMLElement>(selector: string): T => {
  const el = document.querySelector<T>(`.fc-mac-chartdata ${selector}`);
  if (!el) throw new Error(`missing ${selector}`);
  return el;
};
const input = (cls: string): HTMLInputElement => $<HTMLInputElement>(`.${cls} input[type="text"]`);
const change = (el: HTMLInputElement | HTMLSelectElement, value?: string): void => {
  if (value !== undefined) el.value = value;
  el.dispatchEvent(new Event('change', { bubbles: true }));
};
/** Values of the second series of a category chart. */
const secondValues = (chart: ReturnType<typeof sessionChartById>) =>
  chart?.data.domain === 'category' ? chart.data.series[1]?.values : undefined;

const options = (): string[] =>
  Array.from(document.querySelectorAll('.fc-mac-chartdata__option')).map(
    (o) => o.textContent ?? '',
  );

describe('Mac Select Data Source dialog', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;
  let host: HTMLElement;
  let dialog: MacChartSelectDataHandle;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    seed(wb);
    store = createSpreadsheetStore();
    history = new History();
    host = document.createElement('div');
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    dialog = attachMacChartSelectDataDialog({
      host,
      store,
      history,
      getWb: () => wb,
      getLabels: () =>
        ({
          ...SETTINGS_DEFAULT_LABELS,
          ...DEFAULT_SELECT_DATA_LABELS,
          ...ja.sessionCharts,
        }) as SelectDataLabels,
    });
  });

  afterEach(() => {
    dialog.detach();
    document.body.replaceChildren();
    wb.dispose();
  });

  const chart = () => {
    const created = createSessionChart(store, range(0, 0, 3, 2), { wb });
    if (!created) throw new Error('chart was not created');
    return created;
  };

  it('shows the chart range, its series and their cells, with no host feature disabled', () => {
    const created = chart();
    expect(dialog.open(created.id)).toBe(true);
    expect(document.querySelector('.fc-mac-chartdata')?.textContent).toContain(
      'データ ソースの選択',
    );
    expect(input('fc-mac-chartdata__range').value).toBe('=Sheet1!$A$1:$C$4');
    expect(options()).toEqual(['North', 'South']);
    expect(input('fc-mac-chartdata__name').value).toBe('=Sheet1!$B$1');
    expect(input('fc-mac-chartdata__values').value).toBe('=Sheet1!$B$2:$B$4');
    expect(input('fc-mac-chartdata__categories').value).toBe('=Sheet1!$A$2:$A$4');
    expect($<HTMLSelectElement>('.fc-mac-chartdata__empty-select').value).toBe('gap');
    expect($<HTMLInputElement>('.fc-mac-chartdata__show-hidden input').checked).toBe(false);
    const na = $<HTMLInputElement>('.fc-mac-chartdata__na input');
    expect([na.checked, na.disabled]).toEqual([true, true]);
    expect($<HTMLButtonElement>('.fc-mac-chartdata__move-up').disabled).toBe(true);
    expect($<HTMLButtonElement>('.fc-mac-chartdata__move-down').disabled).toBe(true);
    expect($<HTMLButtonElement>('.fc-mac-chartdata__switch').textContent).toBe('行/列の切り替え');
  });

  it('keeps every edit as one undo step on OK', () => {
    const created = chart();
    dialog.open(created.id);
    document.querySelectorAll<HTMLElement>('.fc-mac-chartdata__option')[1]?.click();
    change(input('fc-mac-chartdata__values'), '=Sheet1!$D$2:$D$4');
    change($<HTMLSelectElement>('.fc-mac-chartdata__empty-select'), 'zero');
    const hidden = $<HTMLInputElement>('.fc-mac-chartdata__show-hidden input');
    hidden.checked = true;
    change(hidden);
    $<HTMLButtonElement>('.fc-mac-chartdata__btn--primary').click();

    const after = sessionChartById(store.getState(), created.id);
    expect(secondValues(after)).toEqual([30, 35, 32]);
    expect(after?.emptyCells).toBe('zero');
    expect(after?.showHiddenData).toBe(true);

    expect(history.undo()).toBe(true);
    const back = sessionChartById(store.getState(), created.id);
    expect(secondValues(back)).toEqual([20, 25, 22]);
    expect(back?.emptyCells).toBe('gap');
    expect(back?.showHiddenData).toBe(false);
    expect(history.canUndo()).toBe(false);
  });

  it('rolls every edit back on Cancel and leaves no undo step', () => {
    const created = chart();
    dialog.open(created.id);
    $<HTMLButtonElement>('.fc-mac-chartdata__switch').click();
    expect(options()).toEqual(['Jan', 'Feb', 'Mar']);
    $<HTMLButtonElement>('.fc-mac-chartdata__remove').click();
    expect(options()).toEqual(['Feb', 'Mar']);
    const cancel = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-mac-chartdata__btn'),
    ).find((b) => !b.classList.contains('fc-mac-chartdata__btn--primary'));
    cancel?.click();

    expect(sessionChartById(store.getState(), created.id)?.data.series.map((s) => s.name)).toEqual([
      'North',
      'South',
    ]);
    expect(history.canUndo()).toBe(false);
  });

  it('adds a series once it has values, and reports a bad reference', () => {
    const created = chart();
    dialog.open(created.id);
    $<HTMLButtonElement>('.fc-mac-chartdata__add').click();
    expect(options()).toEqual(['North', 'South', '系列3']);
    change(input('fc-mac-chartdata__values'), 'nonsense');
    expect($('.fc-mac-chartdata__error').hidden).toBe(false);
    change(input('fc-mac-chartdata__name'), '=Sheet1!$D$1');
    change(input('fc-mac-chartdata__values'), '=Sheet1!$D$2:$D$4');
    expect($('.fc-mac-chartdata__error').hidden).toBe(true);
    expect(options()).toEqual(['North', 'South', 'West']);
    $<HTMLButtonElement>('.fc-mac-chartdata__btn--primary').click();
    expect(sessionChartById(store.getState(), created.id)?.data.series).toHaveLength(3);
  });
});
