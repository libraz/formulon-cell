import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { PivotAggregation, PivotAxis, PivotShowValuesAs } from '../../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { ja } from '../../../../src/i18n/strings/ja.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { loadFixture } from '../../engine/pivot-excel-parity/fixture.js';
import { fireContextMenu, item, setRange } from './fixtures.js';

// ja_c0_rows.xlsx: PivotTable1 at A3:C7 on the first sheet. A3 行ラベル, B3/C3 the
// data captions, A4:A6 East/North/South, B4:C6 values, A7:C7 総計.
const CELL = {
  value: [4, 1],
  item: [4, 0],
  header: [2, 0],
  total: [6, 0],
  dataHeader: [2, 1],
} as const;

/** The visible rows of a menu panel: labels, '---' for separators, ✓ for checked. */
const rows = (panel: Element | null): string[] =>
  Array.from(panel?.children ?? []).flatMap((el) => {
    if (el.classList.contains('fc-ctxmenu__sep')) return ['---'];
    if (!el.classList.contains('fc-ctxmenu__item')) return [];
    const label = el.querySelector('.fc-ctxmenu__label')?.textContent ?? '';
    const checked = el.getAttribute('aria-checked') === 'true' ? '✓ ' : '';
    const sub = el.getAttribute('aria-haspopup') === 'menu' ? ' >' : '';
    const disabled = el.getAttribute('aria-disabled') === 'true' ? ' (dis)' : '';
    return [`${checked}${label}${sub}${disabled}`];
  });

describe('PivotTable context menu, as measured on the Mac app', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;
  let detach: ContextMenuHandle | null;
  let fieldListOpened: number;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await loadFixture('ja_c0_rows.xlsx');
    history = new History();
    fieldListOpened = 0;
    detach = attachContextMenu({
      host,
      store,
      wb,
      strings: ja,
      history,
      onOpenPivotFieldList: () => {
        fieldListOpened += 1;
      },
    });
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    document.body.replaceChildren();
    wb.dispose();
  });

  const open = (cell: readonly [number, number]): void => {
    setRange(store, cell[0], cell[1], cell[0], cell[1]);
    fireContextMenu(host, 10, 10);
  };
  const menu = (): string[] => rows(document.querySelector('.fc-ctxmenu'));
  const submenu = (id: string): string[] => {
    const parent = document.querySelector<HTMLButtonElement>(
      `.fc-ctxmenu .fc-ctxmenu__item[data-fc-submenu="${id}"]`,
    );
    if (!parent) throw new Error(`missing submenu ${id}`);
    parent.click();
    return rows(document.querySelector('.fc-ctxmenu__sub'));
  };
  const run = (id: string): void => {
    const button = item(id);
    if (!button) throw new Error(`missing item ${id}`);
    button.click();
  };

  it('lists the value cell menu in the desktop order', () => {
    open(CELL.value);
    expect(menu()).toEqual([
      'コピー',
      'セルの書式設定...',
      '表示形式...',
      '---',
      '更新',
      '---',
      'ピボットテーブルの削除',
      '---',
      '並べ替え >',
      '---',
      '選択 >',
      '数式 >',
      '---',
      '"Summe von Amt" の削除',
      '---',
      '値の集計方法 >',
      '値の表示形式 >',
      '---',
      '詳細の表示 (dis)',
      '---',
      '値フィールドの設定...',
      'ピボットテーブル オプション...',
      'フィールド リストを表示しない',
    ]);
    expect(submenu('pivotSortMenu')).toEqual([
      '昇順',
      '降順',
      '---',
      'その他の並べ替えオプション... (dis)',
    ]);
    expect(submenu('pivotSelectMenu')).toEqual([
      'ラベル (dis)',
      'データ (dis)',
      'ラベルとデータ (dis)',
      'テーブル全体',
      '---',
      '✓ 選択範囲を有効にする (dis)',
    ]);
    expect(submenu('pivotFormulasMenu')).toEqual([
      '集計フィールド... (dis)',
      '集計アイテム... (dis)',
      '---',
      '解決の順序... (dis)',
      '数式の一覧表示 (dis)',
    ]);
    expect(submenu('pivotSummarizeMenu')).toEqual([
      '✓ 合計',
      'データの個数',
      '平均',
      '最大値',
      '最小値',
      '積',
      '個別のカウント (dis)',
      '---',
      'その他のオプション...',
    ]);
    expect(submenu('pivotShowAsMenu')).toEqual([
      '✓ 計算なし',
      '総計に対する比率',
      '列集計に対する比率',
      '行集計に対する比率',
      '親行集計に対する比率',
      '親列集計に対する比率',
      '---',
      'その他のオプション...',
    ]);
  });

  it('lists the row item menu with its field and item names', () => {
    open(CELL.item);
    expect(menu()).toEqual([
      'コピー',
      'セルの書式設定...',
      '---',
      '更新',
      '---',
      'ピボットテーブルの削除',
      '---',
      '並べ替え >',
      'フィルター >',
      '---',
      '選択 >',
      '数式 >',
      '---',
      '✓ "Region" の小計 (dis)',
      '---',
      '展開/折りたたみ >',
      '---',
      'グループ化... (dis)',
      'グループ解除... (dis)',
      '---',
      '移動 >',
      '"Region" の削除',
      '---',
      '値の表示形式 >',
      '---',
      'フィールドの設定...',
      'ピボットテーブル オプション...',
      'フィールド リストを表示しない',
    ]);
    expect(submenu('pivotFilterMenu')).toEqual([
      'フィルターのクリア (dis)',
      '---',
      '選択した項目のみを保持',
      '選択した項目を表示しない',
      '---',
      'トップテン...',
      'ラベル フィルター...',
      '値フィルター...',
    ]);
    expect(submenu('pivotExpandMenu')).toEqual([
      '展開 (dis)',
      '折りたたみ (dis)',
      'フィールド全体の展開 (dis)',
      'フィールド全体の折りたたみ (dis)',
      '---',
      '"Region" の折りたたみ (dis)',
    ]);
    expect(submenu('pivotMoveMenu')).toEqual([
      '"North" を先頭へ移動',
      '"North" を上へ移動',
      '"North" を下へ移動',
      '"North" を末尾へ移動',
      '"Region" を先頭へ移動 (dis)',
      '"Region" を左へ移動 (dis)',
      '"Region" を右へ移動 (dis)',
      '"Region" を末尾へ移動 (dis)',
      '---',
      '"Region" を列に移動',
    ]);
    // The row item draws the whole list disabled.
    expect(submenu('pivotShowAsMenu').every((row) => row === '---' || row.endsWith('(dis)'))).toBe(
      true,
    );
  });

  it('lists the header, grand total and data header menus', () => {
    open(CELL.header);
    expect(menu()).toEqual([
      'コピー',
      'セルの書式設定...',
      '---',
      '更新',
      '---',
      'ピボットテーブルの削除',
      '---',
      '選択 >',
      '数式 >',
      '---',
      '値の表示形式 >',
      '---',
      'ピボットテーブル オプション...',
      'フィールド リストを表示しない',
    ]);
    open(CELL.total);
    expect(menu()).toEqual([
      'コピー',
      'セルの書式設定...',
      '表示形式...',
      '---',
      '更新',
      '---',
      'ピボットテーブルの削除',
      '---',
      '選択 >',
      '数式 >',
      '---',
      '総計の削除',
      '---',
      '値の集計方法 >',
      '値の表示形式 >',
      '---',
      '値フィールドの設定...',
      'ピボットテーブル オプション...',
      'フィールド リストを表示しない',
    ]);
    open(CELL.dataHeader);
    expect(menu()).toEqual([
      'コピー',
      'セルの書式設定...',
      '表示形式...',
      '---',
      '更新',
      '---',
      'ピボットテーブルの削除',
      '---',
      '並べ替え >',
      '---',
      '選択 >',
      '数式 >',
      '---',
      '"Summe von Amt" の削除',
      '値の削除',
      '---',
      '値の集計方法 >',
      '値の表示形式 >',
      '---',
      '値フィールドの設定...',
      'ピボットテーブル オプション...',
      'フィールド リストを表示しない',
    ]);
  });

  it('changes the function from the menu, renaming the caption, as one undo step', () => {
    open(CELL.value);
    submenu('pivotSummarizeMenu');
    run('pivotSummarizeAverage');
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]).toMatchObject({
      aggregation: PivotAggregation.Average,
      name: '平均 / Amt',
    });
    expect(history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]).toMatchObject({
      aggregation: PivotAggregation.Sum,
      name: 'Summe von Amt',
    });
  });

  it('shows values as a share of the grand total from the menu', () => {
    open(CELL.value);
    submenu('pivotShowAsMenu');
    run('pivotShowAsPercentOfTotal');
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]?.showValuesAs).toBe(
      PivotShowValuesAs.PercentOfTotal,
    );
  });

  it('removes a field, moves it to the columns, and deletes the report', () => {
    open(CELL.item);
    submenu('pivotMoveMenu');
    run('pivotMoveFieldToColumns');
    expect(wb.getPivotDefinition(0, 0)?.fields[0]?.axis).toBe(PivotAxis.Col);
    expect(history.undo()).toBe(true);
    expect(wb.getPivotDefinition(0, 0)?.fields[0]?.axis).toBe(PivotAxis.Row);

    open(CELL.item);
    run('pivotRemoveField');
    expect(wb.getPivotDefinition(0, 0)?.fields[0]?.axis).toBeNull();
    expect(history.undo()).toBe(true);

    open(CELL.value);
    run('pivotRemoveField');
    expect(wb.getPivotDefinition(0, 0)?.dataFields.map((d) => d.fieldIndex)).toEqual([3]);
    expect(history.undo()).toBe(true);

    open(CELL.value);
    run('pivotDelete');
    expect(wb.getPivotTables()).toHaveLength(0);
    expect(history.undo()).toBe(true);
    expect(wb.getPivotTables()).toHaveLength(1);
  });

  it('removes the grand total and selects the whole table', () => {
    open(CELL.total);
    run('pivotRemoveGrandTotal');
    expect(wb.getPivotTableGrandTotals(0, 0)).toMatchObject({ cols: false });
    expect(history.undo()).toBe(true);
    expect(wb.getPivotTableGrandTotals(0, 0)).toMatchObject({ cols: true });

    open(CELL.value);
    submenu('pivotSelectMenu');
    run('pivotSelectTable');
    expect(store.getState().selection.range).toMatchObject({ r0: 2, c0: 0, r1: 6, c1: 2 });
  });

  it('opens the field dialog, the options dialog and the field list', () => {
    open(CELL.value);
    run('pivotFieldSettings');
    expect(document.querySelector('.fc-pvfsdlg')?.getAttribute('aria-label')).toBe(
      'ピボットテーブル フィールド',
    );
    document.querySelector<HTMLButtonElement>('.fc-pvfsdlg [data-pvfs-cancel]')?.click();
    open(CELL.header);
    run('pivotOptions');
    expect(document.querySelector('.fc-pvoptdlg')).toBeTruthy();
    document.querySelector<HTMLButtonElement>('.fc-pvoptdlg [data-pvopt-cancel]')?.click();
    open(CELL.header);
    run('pivotShowFieldList');
    expect(fieldListOpened).toBe(1);
  });

  it('filters from the menu: keep or hide the selected items, then clear', () => {
    // The report's row labels are what the filter leaves showing.
    const visible = () =>
      Array.from(wb.pivotCells(0))
        .filter((cell) => cell.addr.col === 0 && cell.addr.row > 2)
        .map((cell) => (cell.value.kind === 'text' ? cell.value.value : ''))
        .filter((label) => label !== '総計');
    open(CELL.item);
    submenu('pivotFilterMenu');
    run('pivotFilterKeepSelected');
    expect(visible()).toEqual(['North']);
    open([3, 0]);
    submenu('pivotFilterMenu');
    expect(item('pivotFilterClear')?.getAttribute('aria-disabled')).not.toBe('true');
    run('pivotFilterClear');
    expect(visible()).toEqual(['East', 'North', 'South']);
    expect(history.undo()).toBe(true);
    expect(visible()).toEqual(['North']);
    expect(history.undo()).toBe(true);
    expect(visible(), 'undo restores every item in its order').toEqual(['East', 'North', 'South']);

    open(CELL.item);
    submenu('pivotFilterMenu');
    run('pivotFilterHideSelected');
    expect(visible()).toEqual(['East', 'South']);
  });

  it('opens the filter panel from Label Filters', () => {
    open(CELL.item);
    submenu('pivotFilterMenu');
    run('pivotFilterLabel');
    expect(document.querySelector('.fc-pivot-filter-panel')).toBeTruthy();
  });

  it('moves an item within its field as one undo step', () => {
    const order = () =>
      Array.from(wb.pivotCells(0))
        .filter((cell) => cell.addr.col === 0 && cell.addr.row > 2)
        .map((cell) => (cell.value.kind === 'text' ? cell.value.value : ''))
        .filter((label) => label !== '総計');
    open(CELL.item);
    submenu('pivotMoveMenu');
    run('pivotMoveItemUp');
    expect(order()).toEqual(['North', 'East', 'South']);
    open([3, 0]);
    submenu('pivotMoveMenu');
    expect(item('pivotMoveItemFirst')?.getAttribute('aria-disabled')).toBe('true');
    run('pivotMoveItemLast');
    expect(order()).toEqual(['East', 'South', 'North']);
    expect(history.undo()).toBe(true);
    expect(order()).toEqual(['North', 'East', 'South']);
    expect(history.undo()).toBe(true);
    expect(order()).toEqual(['East', 'North', 'South']);
  });

  it('builds the menu for the clicked cell, not the active one', async () => {
    const { layoutForView, cellRect } = await import('../../../../src/render/geometry.js');
    setRange(store, CELL.header[0], CELL.header[1], CELL.header[0], CELL.header[1]);
    const s = store.getState();
    const rect = cellRect(layoutForView(s), s.viewport, CELL.total[0], CELL.total[1]);
    fireContextMenu(host, rect.x + 4, rect.y + 4);
    expect(menu()).toContain('総計の削除');
    expect(menu()).toContain('値の集計方法 >');
  });

  it('shows the clicked value cell in the Number Format sample', async () => {
    const { layoutForView, cellRect } = await import('../../../../src/render/geometry.js');
    setRange(store, CELL.header[0], CELL.header[1], CELL.header[0], CELL.header[1]);
    const s = store.getState();
    const rect = cellRect(layoutForView(s), s.viewport, CELL.value[0], CELL.value[1]);
    fireContextMenu(host, rect.x + 4, rect.y + 4);
    run('pivotNumberFormat');
    const preview = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-fmtdlg__preview-cell'),
    ).find((el) => el.closest<HTMLElement>('.fc-fmtdlg')?.hidden === false);
    expect(preview?.textContent).toBe('300');
  });
});

describe('Hide Field List in the PivotTable menu', () => {
  it('is checked while the list is hidden and toggles it', async () => {
    const host = document.createElement('div');
    host.tabIndex = -1;
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const wb = await loadFixture('ja_c0_rows.xlsx');
    let open = true;
    const detach = attachContextMenu({
      host,
      store,
      wb,
      strings: ja,
      onOpenPivotFieldList: () => {
        open = true;
      },
      isPivotFieldListOpen: () => open,
      onClosePivotFieldList: () => {
        open = false;
      },
    });
    const show = (): void => {
      setRange(store, 2, 0, 2, 0);
      fireContextMenu(host, 10, 10);
    };
    show();
    expect(item('pivotShowFieldList')?.getAttribute('aria-checked')).toBe('false');
    item('pivotShowFieldList')?.click();
    expect(open).toBe(false);
    show();
    expect(item('pivotShowFieldList')?.getAttribute('aria-checked')).toBe('true');
    item('pivotShowFieldList')?.click();
    expect(open).toBe(true);
    detach();
    disposeOverlayPortal(host);
    document.body.replaceChildren();
    wb.dispose();
  });
});
