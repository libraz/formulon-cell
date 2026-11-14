import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { openPivotOptions } from '../../../src/interact/pivot-options-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
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

  it('lays out the desktop tabs and options, editable where the report has the option', async () => {
    const wb = await setup();
    const done = openPivotOptions(
      {
        store: createSpreadsheetStore(),
        origin: 'ribbon' as const,
        commandId: 'test',
        host,
        strings: ja,
        wb,
        onChanged: () => {},
      },
      0,
      0,
    );
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
    ).toEqual(['下から上へ', '上から下へ']);
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
    // Only the controls with no report or cache option stay disabled.
    const disabled = Array.from(
      root.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(
        'input, select, textarea',
      ),
    )
      .filter((control) => control.disabled)
      .map(
        (control) => control.closest('label')?.textContent ?? control.getAttribute('aria-label'),
      );
    expect(disabled).toEqual([
      'エラー値に表示する値:',
      '行に表示する',
      '列に表示する',
      '値の領域にフィールドがない場合のアイテムのラベル',
      'ソース データが変更されたときに自動的に更新する',
      'パスワードをファイルに保存する',
      'バックグラウンドでクエリを実行する',
    ]);
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
      {
        store: createSpreadsheetStore(),
        origin: 'ribbon' as const,
        commandId: 'test',
        host,
        strings: ja,
        wb,
        history,
        onChanged: () => (changes += 1),
      },
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

  it('writes the options and the cache options it edits as one undo step, and reopens showing them', async () => {
    const wb = await setup();
    const history = new History();
    let changes = 0;
    const ctx = {
      store: createSpreadsheetStore(),
      origin: 'ribbon' as const,
      commandId: 'test',
      host,
      strings: ja,
      wb,
      history,
      onChanged: () => (changes += 1),
    };
    const before = wb.pivotOptions(0, 0);
    const beforeCache = wb.pivotCacheOptions(0, 0);
    if (!before || !beforeCache) throw new Error('options unreadable');
    const control = <T extends HTMLElement>(selector: string): T => {
      const el = dialog().querySelector<T>(selector);
      if (!el) throw new Error(`missing ${selector}`);
      return el;
    };
    const option = <T extends HTMLElement>(key: string) =>
      control<T>(`[data-pvopt-option="${key}"]`);
    const done = openPivotOptions(ctx, 0, 0);
    option<HTMLInputElement>('showDrill').click();
    option<HTMLInputElement>('mergeItem').click();
    option<HTMLInputElement>('enableDrill').click();
    const errorBox = option<HTMLInputElement>('showError');
    expect(option<HTMLInputElement>('errorCaption').disabled).toBe(!errorBox.checked);
    if (!errorBox.checked) errorBox.click();
    errorBox.dispatchEvent(new Event('change'));
    expect(option<HTMLInputElement>('errorCaption').disabled).toBe(false);
    option<HTMLInputElement>('errorCaption').value = 'n/a';
    const indent = option<HTMLSelectElement>('indent');
    expect(indent.value).toBe(String(before.indent + 1));
    indent.value = '4';
    option<HTMLSelectElement>('pageWrap').value = '3';
    control<HTMLInputElement>('[data-pvopt-option="pageOverThenDown"] input[value="1"]').click();
    option<HTMLInputElement>('altText').value = 'Sales';
    option<HTMLTextAreaElement>('altTextSummary').value = 'Amount by region';
    control<HTMLInputElement>('[data-pvopt-cache="refreshOnLoad"]').click();
    control<HTMLButtonElement>('[data-pvopt-ok]').click();
    expect(await done).toBe(true);
    expect(changes).toBe(1);

    const expected = {
      ...before,
      showDrill: !before.showDrill,
      mergeItem: !before.mergeItem,
      enableDrill: !before.enableDrill,
      showError: true,
      errorCaption: 'n/a',
      indent: 3,
      pageWrap: 3,
      pageOverThenDown: true,
      altText: 'Sales',
      altTextSummary: 'Amount by region',
    };
    expect(wb.pivotOptions(0, 0)).toEqual(expected);
    expect(wb.pivotCacheOptions(0, 0)).toEqual({
      ...beforeCache,
      refreshOnLoad: !beforeCache.refreshOnLoad,
    });

    const reopened = openPivotOptions(ctx, 0, 0);
    expect(option<HTMLInputElement>('showDrill').checked).toBe(expected.showDrill);
    expect(option<HTMLInputElement>('errorCaption').value).toBe('n/a');
    expect(option<HTMLSelectElement>('indent').value).toBe('4');
    expect(option<HTMLSelectElement>('pageWrap').value).toBe('3');
    expect(
      control<HTMLInputElement>('[data-pvopt-option="pageOverThenDown"] input[value="1"]').checked,
    ).toBe(true);
    expect(option<HTMLTextAreaElement>('altTextSummary').value).toBe('Amount by region');
    control<HTMLButtonElement>('[data-pvopt-cancel]').click();
    expect(await reopened).toBe(false);

    expect(history.undo()).toBe(true);
    expect(wb.pivotOptions(0, 0)).toEqual(before);
    expect(wb.pivotCacheOptions(0, 0)).toEqual(beforeCache);
    expect(history.redo()).toBe(true);
    expect(wb.pivotOptions(0, 0)).toEqual(expected);
    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(reloaded.pivotOptions(0, 0)).toEqual(expected);
      expect(reloaded.pivotCacheOptions(0, 0)?.refreshOnLoad).toBe(!beforeCache.refreshOnLoad);
    } finally {
      reloaded.dispose();
    }
    wb.dispose();
  });

  it('writes every wired boolean and caption option and keeps them through save and load', async () => {
    const wb = await setup();
    const history = new History();
    const ctx = {
      store: createSpreadsheetStore(),
      origin: 'ribbon' as const,
      commandId: 'test',
      host,
      strings: ja,
      wb,
      history,
      onChanged: () => {},
    };
    const before = wb.pivotOptions(0, 0);
    const beforeCache = wb.pivotCacheOptions(0, 0);
    if (!before || !beforeCache) throw new Error('options unreadable');
    const flags = [
      'showDrill',
      'printDrill',
      'showDataTips',
      'showHeaders',
      'gridDropZones',
      'showError',
      'showMissing',
      'itemPrintTitles',
      'fieldPrintTitles',
      'multipleFieldFilters',
      'autoFormatting',
      'preserveFormatting',
      'mergeItem',
      'customListSort',
      'enableDrill',
    ] as const;
    const done = openPivotOptions(ctx, 0, 0);
    for (const key of flags) {
      document
        .querySelector<HTMLInputElement>(`.fc-pvoptdlg [data-pvopt-option="${key}"]`)
        ?.click();
    }
    const missing = document.querySelector<HTMLInputElement>(
      '.fc-pvoptdlg [data-pvopt-option="missingCaption"]',
    );
    if (!missing) throw new Error('missing caption field absent');
    missing.value = 'none';
    document.querySelector<HTMLInputElement>('.fc-pvoptdlg [data-pvopt-cache="saveData"]')?.click();
    document.querySelector<HTMLButtonElement>('.fc-pvoptdlg [data-pvopt-ok]')?.click();
    expect(await done).toBe(true);

    const expected = {
      ...before,
      ...Object.fromEntries(flags.map((key) => [key, !before[key]])),
      missingCaption: 'none',
    };
    expect(wb.pivotOptions(0, 0)).toEqual(expected);
    expect(wb.pivotCacheOptions(0, 0)?.saveData).toBe(!beforeCache.saveData);
    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(reloaded.pivotOptions(0, 0)).toEqual(expected);
    } finally {
      reloaded.dispose();
    }
    expect(history.undo()).toBe(true);
    expect(wb.pivotOptions(0, 0)).toEqual(before);
    expect(wb.pivotCacheOptions(0, 0)).toEqual(beforeCache);
    wb.dispose();
  });

  it('renders the controls with no engine option disabled through the shared projection', async () => {
    const wb = await setup();
    const done = openPivotOptions(
      {
        store: createSpreadsheetStore(),
        origin: 'ribbon' as const,
        commandId: 'test',
        host,
        strings: ja,
        wb,
        history: new History(),
        onChanged: () => {},
      },
      0,
      0,
    );
    const unbound = Array.from(
      dialog().querySelectorAll<HTMLInputElement>('input[type="checkbox"], input[type="radio"]'),
    ).filter(
      (el) =>
        !el.dataset.pvoptOption &&
        !el.dataset.pvoptCache &&
        !el.closest('[data-pvopt-option], [data-pvopt-limit]'),
    );
    expect(unbound.length).toBe(6);
    for (const el of unbound) {
      expect(el.disabled).toBe(true);
      expect(el.dataset.disabledReason).toBe(ja.mac.pivotOptions.unavailable);
    }
    expect(unbound.filter((el) => el.type === 'checkbox').map((el) => el.checked)).toEqual([
      false,
      false,
      true,
      false,
      false,
      false,
    ]);
    dialog().querySelector<HTMLButtonElement>('[data-pvopt-cancel]')?.click();
    expect(await done).toBe(false);
    wb.dispose();
  });

  describe('retain items deleted from the data source', () => {
    const open = async (limit: number | null) => {
      const wb = await setup();
      if (limit !== null)
        expect(wb.updatePivotCacheOptions(0, 0, { missingItemsLimit: limit })).toBe(true);
      const history = new History();
      const done = openPivotOptions(
        {
          store: createSpreadsheetStore(),
          origin: 'ribbon' as const,
          commandId: 'test',
          host,
          strings: ja,
          wb,
          history,
          onChanged: () => {},
        },
        0,
        0,
      );
      const box = dialog().querySelector<HTMLInputElement>(
        '[data-pvopt-cache="missingItemsLimit"]',
      );
      const radios = Array.from(
        dialog().querySelectorAll<HTMLInputElement>('[data-pvopt-limit] input'),
      );
      if (!box || radios.length !== 2) throw new Error('retain controls missing');
      const click = (selector: string) =>
        dialog().querySelector<HTMLButtonElement>(selector)?.click();
      return {
        wb,
        box,
        radios,
        done,
        history,
        ok: () => click('[data-pvopt-ok]'),
        cancel: () => click('[data-pvopt-cancel]'),
        toggle: () => box.click(),
      };
    };
    type Opened = Awaited<ReturnType<typeof open>>;
    const shown = (radios: HTMLInputElement[]) => radios.findIndex((r) => r.checked);

    it.each([
      [null, true, 0, false],
      [0, false, 1, true],
      [32500, true, 1, false],
      [50, true, 1, false],
    ])(
      'shows %s as checked=%s, radio %s, radios disabled=%s',
      async (limit, checked, radio, off) => {
        const { wb, box, radios, cancel, done } = await open(limit);
        expect(box.checked).toBe(checked);
        expect(shown(radios)).toBe(radio);
        expect(radios.map((r) => r.disabled)).toEqual([off, off]);
        expect(radios.map((r) => r.parentElement?.textContent)).toEqual(['自動', '最大値']);
        cancel();
        expect(await done).toBe(false);
        wb.dispose();
      },
    );

    it.each<[string, number | null, number | null, (c: Opened) => void]>([
      ['unchecked', null, 0, (c) => c.toggle()],
      ['checked + Max', 0, 1048576, (c) => c.toggle()],
      [
        'checked + Automatic',
        0,
        null,
        (c) => {
          c.toggle();
          c.radios[0]?.click();
        },
      ],
      ['32500 unchanged', 32500, 1048576, () => {}],
      ['50 unchanged', 50, 1048576, () => {}],
      ['null unchanged', null, null, () => {}],
      ['0 unchanged', 0, 0, () => {}],
    ])('OK writes %s', async (_name, from, to, act) => {
      const c = await open(from);
      act(c);
      c.ok();
      expect(await c.done).toBe(true);
      expect(c.wb.pivotCacheOptions(0, 0)?.missingItemsLimit).toBe(to);
      c.wb.dispose();
    });

    it('checking the box on a 0 book selects Max and enables the radios', async () => {
      const c = await open(0);
      c.toggle();
      expect(shown(c.radios)).toBe(1);
      expect(c.radios.map((r) => r.disabled)).toEqual([false, false]);
      c.cancel();
      await c.done;
      c.wb.dispose();
    });

    it('unchecking shows Max on the disabled radios, and undo restores the value', async () => {
      const c = await open(null);
      c.toggle();
      expect(shown(c.radios)).toBe(1);
      expect(c.radios.map((r) => r.disabled)).toEqual([true, true]);
      c.ok();
      await c.done;
      expect(c.history.undo()).toBe(true);
      expect(c.wb.pivotCacheOptions(0, 0)?.missingItemsLimit).toBe(null);
      c.wb.dispose();
    });

    it('Cancel keeps 50', async () => {
      const c = await open(50);
      c.toggle();
      c.toggle();
      c.radios[0]?.click();
      c.cancel();
      expect(await c.done).toBe(false);
      expect(c.wb.pivotCacheOptions(0, 0)?.missingItemsLimit).toBe(50);
      c.wb.dispose();
    });
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
