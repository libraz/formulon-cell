import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { findPivotTableAtCell } from '../../../src/engine/passthrough-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachPivotTableDialog } from '../../../src/interact/pivot-table-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const buildWorkbook = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  const rows: [string, string, number][] = [
    ['Region', 'Product', 0],
    ['East', 'A', 500],
    ['North', 'A', 100],
    ['North', 'B', 200],
    ['South', 'A', 300],
  ];
  rows.forEach(([region, product, amt], row) => {
    wb.setText({ sheet: 0, row, col: 0 }, region);
    wb.setText({ sheet: 0, row, col: 1 }, product);
    if (row === 0) wb.setText({ sheet: 0, row, col: 2 }, 'Amt');
    else wb.setNumber({ sheet: 0, row, col: 2 }, amt);
  });
  return wb;
};

// Excel for Mac (ja): source and destination only; OK creates an empty report.
describe('Create PivotTable dialog on the Mac host', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.replaceChildren();
  });

  it('shows the measured source and destination controls only', async () => {
    const wb = await buildWorkbook();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: ja });
    try {
      handle.open();
      const panel = document.querySelector('.fc-pivotdlg--mac .fc-pivotdlg__panel');
      expect(panel).toBeTruthy();
      const labels = Array.from(
        document.querySelectorAll(
          '.fc-pivotdlg__placement-label, .fc-pivotdlg__field > span:first-child',
        ),
      ).map((el) => el.textContent);
      expect(labels).toEqual([
        '分析するデータを選択してください。',
        'テーブル/範囲:',
        'ピボットテーブルの作成先を選択してください。',
        'テーブル/範囲:',
      ]);
      const radios = Array.from(
        document.querySelectorAll<HTMLInputElement>('.fc-pivotdlg__check input[type="radio"]'),
      ).map((input) => [input.parentElement?.textContent, input.checked, input.disabled]);
      expect(radios).toEqual([
        ['テーブルまたは範囲を選択', true, false],
        ['外部データ ソースを使用', false, false],
        ['新規ワークシート', true, false],
        ['既存のワークシート', false, false],
      ]);
      const connection = document.querySelector<HTMLButtonElement>('.fc-pivotdlg__connection-btn');
      expect(connection?.textContent).toBe('接続の選択...');
      expect(connection?.disabled).toBe(true);
      expect(connection?.nextElementSibling?.textContent).toBe('取り出されたデータはありません。');
      const inputs = document.querySelectorAll<HTMLInputElement>('.fc-range-picker input');
      expect(Array.from(inputs, (input) => input.value)).toEqual(['Sheet1!$A$1:$C$5', '']);
      expect(document.activeElement).toBe(inputs[0]);
      expect(document.querySelector('.fc-pivotdlg__field-list')).toBeNull();
      expect(document.querySelector('.fc-pivotdlg__checkgrid')).toBeNull();
      expect(document.querySelectorAll('.fc-pivotdlg select')).toHaveLength(0);
      expect(
        Array.from(document.querySelectorAll('.fc-pivotdlg button'))
          .map((b) => b.textContent)
          .slice(-2),
      ).toEqual(['キャンセル', 'OK']);
    } finally {
      handle.detach();
      wb.dispose();
    }
  });

  it('lays each range field out inline with its picker inside the field', async () => {
    const wb = await buildWorkbook();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: ja });
    try {
      handle.open();
      const rows = Array.from(document.querySelectorAll<HTMLElement>('.fc-pivotdlg__field'));
      expect(rows).toHaveLength(2);
      for (const row of rows) {
        expect(row.classList.contains('fc-pivotdlg__field--inline')).toBe(true);
        expect(row.children[0]?.textContent).toBe('テーブル/範囲:');
        const wrap = row.children[1] as HTMLElement;
        expect(wrap.classList.contains('fc-range-picker')).toBe(true);
        expect(wrap.querySelector('input')).toBeTruthy();
        expect(wrap.querySelector('.fc-range-picker__btn')).toBeTruthy();
        expect(wrap.querySelector('input')?.placeholder).toBe('');
      }
    } finally {
      handle.detach();
      wb.dispose();
    }
  });

  it('disables OK while the external-data choice has no connection', async () => {
    const wb = await buildWorkbook();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: ja });
    try {
      handle.open();
      const [range, external] = Array.from(
        document.querySelectorAll<HTMLInputElement>('input[name="fc-pivotdlg-source-kind"]'),
      );
      const ok = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-pivotdlg button')).at(
        -1,
      );
      expect(ok?.disabled).toBe(false);
      external?.click();
      expect(ok?.disabled).toBe(true);
      range?.click();
      expect(ok?.disabled).toBe(false);
    } finally {
      handle.detach();
      wb.dispose();
    }
  });

  it('creates an empty report on a new sheet and selects it', async () => {
    const wb = await buildWorkbook();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: ja });
    try {
      handle.open();
      document
        .querySelector('form')
        ?.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));
      const active = store.getState().selection.active;
      expect(active).toEqual({ sheet: 0, row: 2, col: 0 });
      const pivot = findPivotTableAtCell(wb, active);
      expect(pivot).not.toBeNull();
      const def = wb.getPivotDefinition(active.sheet, pivot?.pivotIndex ?? -1);
      expect(def?.dataFields).toEqual([]);
      expect(def?.fields.map((f) => [f.sourceName, f.axis])).toEqual([
        ['Region', null],
        ['Product', null],
        ['Amt', null],
      ]);
    } finally {
      handle.detach();
      wb.dispose();
    }
  });
});
