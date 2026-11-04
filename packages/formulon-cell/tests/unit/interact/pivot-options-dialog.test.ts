import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { openPivotOptions } from '../../../src/interact/pivot-options-dialog.js';
import { mutators } from '../../../src/store/store.js';
import { mountStubSheet } from '../../test-utils/index.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Amt'],
    ['East', 12],
    ['West', 5],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  expect(
    createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      destination: { sheet: 0, row: 6, col: 0 },
      name: 'PivotTable1',
      rowField: 'Region',
      valueField: 'Amt',
    }).ok,
  ).toBe(true);
  return wb;
};

describe('PivotTable Options dialog', () => {
  let host: HTMLElement;
  beforeEach(() => {
    document.body.replaceChildren();
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => document.body.replaceChildren());

  const dialog = (): HTMLElement => {
    const el = document.querySelector<HTMLElement>('.fc-pvoptdlg');
    if (!el) throw new Error('dialog not open');
    return el;
  };
  const panel = (id: string): HTMLElement => {
    const el = dialog().querySelector<HTMLElement>(`[data-pvopt-panel="${id}"]`);
    if (!el) throw new Error(`missing ${id} panel`);
    return el;
  };
  const checks = (id: string) =>
    Array.from(panel(id).querySelectorAll<HTMLLabelElement>('.fc-pvoptdlg__check')).map((label) => [
      label.textContent,
      label.querySelector('input')?.checked,
    ]);

  it('lays out the desktop tabs and options, with only the name editable', async () => {
    const wb = await setup();
    const done = openPivotOptions({ host, strings: ja, wb, onChanged: () => {} }, 0, 0);
    const root = dialog();
    expect(root.getAttribute('aria-label')).toBe('ピボットテーブル オプション');
    expect(root.querySelector('.fc-pvoptdlg__name-row')?.textContent).toBe('ピボットテーブル名:');
    expect(root.querySelector<HTMLInputElement>('[data-pvopt-name]')?.value).toBe('PivotTable1');
    const nameInput = root.querySelector<HTMLInputElement>('[data-pvopt-name]');
    expect(document.activeElement).toBe(nameInput);
    expect([nameInput?.selectionStart, nameInput?.selectionEnd]).toEqual([0, 'PivotTable1'.length]);
    expect(
      Array.from(root.querySelectorAll('[role="tab"]')).map((tab) => [
        tab.textContent,
        tab.getAttribute('aria-selected'),
      ]),
    ).toEqual([
      ['表示形式', 'true'],
      ['レイアウト', 'false'],
      ['データ', 'false'],
      ['代替テキスト', 'false'],
    ]);
    expect(
      Array.from(panel('display').querySelectorAll('.fc-pvoptdlg__group-label')).map(
        (label) => label.textContent,
      ),
    ).toEqual(['表示:', 'データのないアイテム:', '印刷:']);
    expect(checks('display')).toEqual([
      ['展開/折りたたみボタン', true],
      ['展開/折りたたみボタンをプリントする', false],
      ['操作に応じたヒント', true],
      ['フィールドのタイトルとフィルター', true],
      ['従来のピボットテーブル レイアウト', false],
      ['エラー値に表示する値:', false],
      ['空白セルに表示する値:', true],
      ['行に表示する', false],
      ['列に表示する', false],
      ['各印刷ページに行ラベルを繰り返す', true],
      ['印刷タイトルを設定する', false],
      ['値の領域にフィールドがない場合のアイテムのラベル', true],
    ]);
    expect(checks('layout').map(([label]) => label)).toEqual([
      '1 つのフィールドに複数のフィルターを使用可能にする',
      '更新時に列幅を自動調整する',
      '更新時にセル書式を保持する',
      'セルとラベルを結合して中央揃えにする',
      '並べ替え時にユーザー設定リストを使用する',
    ]);
    expect(
      Array.from(panel('layout').querySelectorAll('.fc-pvoptdlg__radio')).map((r) => r.textContent),
    ).toEqual(['下から上へ', '上から下へ', '昇順で並べ替え', 'データ ソース順で並べ替え']);
    expect(checks('data').map(([label]) => label)).toEqual([
      'ファイルに元のデータを保存する',
      '詳細を表示可能にする',
      'データ ソースから削除されたアイテムの保持',
      'ファイルを開くときにデータを更新する',
      'ソース データが変更されたときに自動的に更新する',
      'パスワードをファイルに保存する',
      'バックグラウンドでクエリを実行する',
    ]);
    expect(panel('altText').querySelector('.fc-pvoptdlg__alt-help')?.textContent).toMatch(
      /^タイトルと説明は、表、図表、画像/,
    );
    const editable = Array.from(
      root.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(
        'input, select, textarea',
      ),
    ).filter((control) => !control.disabled);
    expect(editable.map((control) => control.dataset.pvoptName !== undefined)).toEqual([true]);
    expect(
      Array.from(root.querySelectorAll('.fc-pvoptdlg__footer button')).map((b) => b.textContent),
    ).toEqual(['キャンセル', 'OK']);
    root.querySelector<HTMLButtonElement>('[data-pvopt-cancel]')?.click();
    expect(await done).toBe(false);
    wb.dispose();
  });

  it('renames the report as one undo step', async () => {
    const wb = await setup();
    const history = new History();
    let changes = 0;
    const done = openPivotOptions(
      { host, strings: ja, wb, history, onChanged: () => (changes += 1) },
      0,
      0,
    );
    const name = dialog().querySelector<HTMLInputElement>('[data-pvopt-name]');
    if (!name) throw new Error('missing name');
    name.value = 'Sales by region';
    dialog().querySelector<HTMLButtonElement>('[data-pvopt-ok]')?.click();
    expect(await done).toBe(true);
    expect(wb.pivotTableName(0, 0)).toBe('Sales by region');
    expect(changes).toBe(1);
    expect(history.undo()).toBe(true);
    expect(wb.pivotTableName(0, 0)).toBe('PivotTable1');
    expect(history.redo()).toBe(true);
    expect(wb.pivotTableName(0, 0)).toBe('Sales by region');
    wb.dispose();
  });

  it('opens from the instance on the active cell report, and renames through it', async () => {
    const wb = await setup();
    const sheet = await mountStubSheet({ workbook: wb });
    const { instance } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 5 });
    expect(instance.openPivotOptions()).toBe(false);
    mutators.setActive(instance.store, { sheet: 0, row: 7, col: 0 });
    expect(instance.openPivotOptions()).toBe(true);
    const name = document.querySelector<HTMLInputElement>('.fc-pvoptdlg [data-pvopt-name]');
    if (!name) throw new Error('dialog not open');
    name.value = 'Regions';
    document.querySelector<HTMLButtonElement>('.fc-pvoptdlg [data-pvopt-ok]')?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(wb.pivotTableName(0, 0)).toBe('Regions');
    expect(instance.openPivotOptions(0, 3)).toBe(false);
    sheet.dispose();
  });
});
