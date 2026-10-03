import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { dictionaries } from '../../../../src/i18n/strings.js';
import { createBordersMenu } from '../../../../src/toolbar/ribbon/menus/borders.js';
import { createConditionalMenu } from '../../../../src/toolbar/ribbon/menus/conditional.js';
import { createFormulasMenuFactories } from '../../../../src/toolbar/ribbon/menus/formulas.js';
import { createHomeMenuFactories } from '../../../../src/toolbar/ribbon/menus/home.js';
import { createStylesMenuFactories } from '../../../../src/toolbar/ribbon/menus/styles.js';
import { readCssWithImports } from '../../../helpers/css.js';
import { root } from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
  it('keeps Underline dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createUnderlineMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-underline-action]'));

    expect(menu.id).toBe('menu-underline');
    expect(items.map((item) => item.dataset.underlineAction)).toEqual(['single', 'double']);
    expect(items.map((item) => item.textContent)).toEqual(['下線', '二重下線']);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(2);

    expect(menusCss).toMatch(/#menu-underline\s*\{[\s\S]*?min-width: 118px;/);
    expect(menusCss).toMatch(
      /#menu-underline \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-underline \.fc-tb__menu-icon,[\s\S]*?#menu-underline \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Copy dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createCopyMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-copy-action]'));

    expect(menu.id).toBe('menu-copy');
    expect(items.map((item) => item.dataset.copyAction)).toEqual(['copy', 'picture']);
    expect(items.map((item) => item.textContent)).toEqual(['コピー', '図としてコピー...']);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(2);

    expect(menusCss).toMatch(/#menu-copy\s*\{[\s\S]*?min-width: 144px;/);
    expect(menusCss).toMatch(
      /#menu-copy \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-copy \.fc-tb__menu-icon,[\s\S]*?#menu-copy \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Paste dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(/#menu-paste\s*\{[\s\S]*?min-width: 198px;/);
    expect(menusCss).toMatch(
      /#menu-paste \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-paste \.fc-tb__menu-icon,[\s\S]*?#menu-paste \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Clear dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createClearMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-clear]'));

    expect(menu.id).toBe('menu-clear');
    expect(items.map((item) => item.dataset.clear)).toEqual([
      'all',
      'formats',
      'contents',
      'comments',
      'hyperlinks',
      'remove-hyperlinks',
      'conditional',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      'すべてクリア',
      '書式のクリア',
      '数式と値のクリア',
      'コメントとメモのクリア',
      'ハイパーリンクのクリア',
      'ハイパーリンクの削除',
      '条件付き書式のクリア',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(7);
    expect(menu.querySelector('path[fill="#f7e1ff"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    // The destructive mark is a filled badge plate, not a stroked cross.
    expect(menu.querySelector('path[fill="#c00000"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-clear\s*\{[\s\S]*?min-width: 194px;/);
    expect(menusCss).toMatch(
      /#menu-clear \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-clear \.fc-tb__menu-icon,[\s\S]*?#menu-clear \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Sort and Filter dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createSortMenu('sortFilterHome');
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-sort]'));

    expect(menu.id).toBe('menu-sort-home');
    expect(items.map((item) => item.dataset.sort)).toEqual([
      'asc',
      'desc',
      'custom',
      'filter',
      'filter-by-value',
      'filter-clear',
      'filter-reapply',
      'filter-advanced',
      'dedupe',
      'conditional',
      'named',
    ]);
    expect(items.slice(0, 8).map((item) => item.textContent)).toEqual([
      '昇順で並べ替え',
      '降順で並べ替え',
      'ユーザー設定の並べ替え...',
      'フィルター',
      '選択したセルの値でフィルター',
      'クリア',
      '再適用',
      '詳細設定...',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(11);
    expect(menu.querySelector('path[stroke="#c00000"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-sort-home,[\s\S]*?#menu-sort\s*\{[\s\S]*?min-width: 218px;/);
    expect(menusCss).toMatch(
      /#menu-sort-home \.fc-tb__menu-item,[\s\S]*?#menu-sort \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-sort-home \.fc-tb__menu-icon,[\s\S]*?#menu-sort \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Find and Select dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createFindSelectMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-find-select]'));

    expect(menu.id).toBe('menu-find-select');
    expect(items.map((item) => item.dataset.findSelect)).toEqual([
      'find',
      'replace',
      'go-to',
      'go-to-special',
      'formulas',
      'comments',
      'conditional-format',
      'constants',
      'data-validation',
      'object-select',
      'selection-pane',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '検索...',
      '置換...',
      'ジャンプ...',
      '条件を選択してジャンプ...',
      '数式',
      'コメントとメモ',
      '条件付き書式',
      '定数',
      'データの入力規則',
      'オブジェクトの選択',
      '選択ウィンドウ...',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(11);
    expect(menu.querySelector('path[stroke="#8a4fb7"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(menu.querySelector('path[fill="#fdf3bf"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-find-select\s*\{[\s\S]*?min-width: 182px;/);
    expect(menusCss).toMatch(
      /#menu-find-select \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-find-select \.fc-tb__menu-icon,[\s\S]*?#menu-find-select \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Fill dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createFillMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-fill]'));

    expect(menu.id).toBe('menu-fill');
    expect(items.map((item) => item.dataset.fill)).toEqual([
      'down',
      'right',
      'up',
      'left',
      'group',
      'series',
      'justify',
      'flash',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '下方向へコピー',
      '右方向へコピー',
      '上方向へコピー',
      '左方向へコピー',
      '作業グループへコピー...',
      '連続データの作成...',
      '文字の割付',
      'フラッシュ フィル',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(8);
    expect(menu.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(menu.querySelector('path[fill="#ed7d31"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-fill\s*\{[\s\S]*?min-width: 178px;/);
    expect(menusCss).toMatch(
      /#menu-fill \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-fill \.fc-tb__menu-icon,[\s\S]*?#menu-fill \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps AutoSum dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const factories = createFormulasMenuFactories(ja.ribbonMenu, 'ja');
    const homeMenu = factories.createAutoSumMenu('autosum');
    const formulasMenu = factories.createAutoSumMenu('autosumFormula');
    const items = Array.from(homeMenu.querySelectorAll<HTMLButtonElement>('[data-autosum-fn]'));

    expect(homeMenu.id).toBe('menu-autosum-home');
    expect(formulasMenu.id).toBe('menu-autosum-formulas');
    expect(items.map((item) => item.dataset.autosumFn)).toEqual([
      'SUM',
      'AVERAGE',
      'COUNT',
      'MAX',
      'MIN',
      'MORE',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '合計',
      '平均',
      '数値の個数',
      '最大値',
      '最小値',
      'その他の関数...',
    ]);
    expect(homeMenu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(1);
    expect(homeMenu.querySelector('[data-autosum-fn="SUM"] .fc-tb__menu-icon-svg')).toBeTruthy();
    expect(homeMenu.querySelectorAll('.fc-tb__menu-item__icon-spacer')).toHaveLength(5);

    expect(menusCss).toMatch(
      /#menu-autosum-home,[\s\S]*?#menu-autosum-formulas\s*\{[\s\S]*?min-width: 128px;/,
    );
    expect(menusCss).toMatch(
      /#menu-autosum-home \.fc-tb__menu-item,[\s\S]*?#menu-autosum-formulas \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).not.toContain('.fc-tb__menu-icon--autosum-average::before');
    expect(menusCss).not.toContain('.fc-tb__menu-icon--autosum-more::after');
  });

  it('keeps Currency dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createStylesMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
    }).createCurrencyMenu();
    const presetItems = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('[data-currency-preset]'),
    );
    const footer = menu.querySelector<HTMLButtonElement>('[data-currency-footer]');

    expect(menu.id).toBe('menu-currency-home');
    expect(presetItems.map((item) => item.dataset.currencyPreset)).toEqual([
      '¥',
      '$',
      '€',
      '£',
      'CHF',
    ]);
    expect(presetItems.map((item) => item.textContent)).toEqual([
      '¥ 日本語',
      '$ 英語 (米国)',
      '€ ユーロ (€ 123)',
      '£ 英語 (英国)',
      'CHF フランス語 (スイス)',
    ]);
    expect(footer?.textContent).toBe('その他の通貨表示形式…');
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(0);
    expect(menu.querySelectorAll('.fc-tb__menu-item__icon-spacer')).toHaveLength(6);

    expect(menusCss).toMatch(/\.fc-tb__currency-menu\s*\{[\s\S]*?min-width: 190px;/);
    expect(menusCss).toMatch(
      /\.fc-tb__currency-menu \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__currency-menu \.fc-tb__menu-item__icon-spacer\s*\{[\s\S]*?width: 0;/,
    );
    expect(menusCss).not.toContain('.fc-tb__menu-icon--currency-yen::before');
    expect(menusCss).not.toContain('.fc-tb__menu-icon--currency-more::before');
  });

  it('keeps Cell Styles gallery geometry close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createStylesMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
    }).createCellStylesMenu();
    const scrollBody = menu.querySelector<HTMLElement>(':scope > .fc-tb__cellstyle-scroll');
    const headings = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__cellstyle-heading') ?? [],
    ).map((heading) => heading.textContent);
    const grids = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__cellstyle-grid') ?? [],
    );

    expect(menu.id).toBe('menu-cell-styles-home');
    expect(menu.classList.contains('fc-tb__cellstyle-menu')).toBe(true);
    expect(headings).toEqual([
      '良い、悪い、標準',
      'データとモデル',
      'タイトルと見出し',
      'テーマのセル スタイル',
      '表示形式',
    ]);
    expect(grids.length).toBeGreaterThanOrEqual(5);
    expect(menu.querySelector('[data-cell-style="normal"]')?.textContent).toBe('標準');
    expect(menu.querySelector('[data-cell-style="good"]')?.textContent).toBe('良い');
    expect(menu.querySelectorAll('.fc-tb__cellstyle-footer')).toHaveLength(2);

    expect(menusCss).toMatch(/\.fc-tb__cellstyle-menu\s*\{[\s\S]*?width: 566px;/);
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-scroll\s*\{[\s\S]*?max-height: min\(388px, calc\(100vh - 244px\)\);/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-grid\s*\{[\s\S]*?grid-template-columns: repeat\(6, 82px\);[\s\S]*?gap: 7px 11px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-heading\s*\{[\s\S]*?background: #f3f2f1;[\s\S]*?font-weight: 400;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-chip\s*\{[\s\S]*?min-width: 82px;[\s\S]*?height: 22px;[\s\S]*?border-radius: 0;/,
    );
  });

  it('keeps Format as Table gallery geometry close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createStylesMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
    }).createTableStyleMenu('formatTableHome');
    const scrollBody = menu.querySelector<HTMLElement>(':scope > .fc-tb__tablestyle-scroll');
    const headings = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__tablestyle-heading') ?? [],
    ).map((heading) => heading.textContent);
    const grids = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__tablestyle-grid') ?? [],
    );

    expect(menu.id).toBe('menu-table-style-home');
    expect(menu.classList.contains('fc-tb__tablestyle-menu')).toBe(true);
    expect(headings).toEqual(['淡色', '中間', '濃色']);
    expect(grids.map((grid) => grid.querySelectorAll('.fc-tb__tablestyle-swatch').length)).toEqual([
      28, 28, 7,
    ]);
    expect(menu.querySelectorAll('.fc-tb__tablestyle-footer')).toHaveLength(2);

    expect(menusCss).toMatch(/\.fc-tb__tablestyle-menu\s*\{[\s\S]*?width: 515px;/);
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-scroll\s*\{[\s\S]*?max-height: min\(496px, calc\(100vh - 116px\)\);/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-grid\s*\{[\s\S]*?grid-template-columns: repeat\(7, 62px\);[\s\S]*?gap: 8px 9px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-heading\s*\{[\s\S]*?background: #f3f2f1;[\s\S]*?font-weight: 400;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-swatch\s*\{[\s\S]*?width: 62px;[\s\S]*?height: 47px;[\s\S]*?border-radius: 0;/,
    );
  });

  it('keeps Insert and Delete Cells dropdowns close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const factories = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    });
    const insertMenu = factories.createInsertCellsMenu();
    const deleteMenu = factories.createDeleteCellsMenu();
    const insertItems = Array.from(
      insertMenu.querySelectorAll<HTMLButtonElement>('[data-cell-insert]'),
    );
    const deleteItems = Array.from(
      deleteMenu.querySelectorAll<HTMLButtonElement>('[data-cell-delete]'),
    );

    expect(insertMenu.id).toBe('menu-insert-cells');
    expect(insertItems.map((item) => item.dataset.cellInsert)).toEqual([
      'cells',
      'rows',
      'cols',
      'sheet',
    ]);
    expect(insertItems.map((item) => item.textContent)).toEqual([
      'セルを挿入...',
      'シートの行を挿入',
      'シートの列を挿入',
      'シートの挿入',
    ]);
    expect(deleteMenu.id).toBe('menu-delete-cells');
    expect(deleteItems.map((item) => item.dataset.cellDelete)).toEqual([
      'cells',
      'rows',
      'cols',
      'row',
      'col',
      'sheet',
    ]);
    expect(deleteItems.map((item) => item.textContent)).toEqual([
      'セルを削除...',
      'シートの行を削除',
      'シートの列を削除',
      '行の削除',
      '列の削除',
      'シートの削除',
    ]);
    expect(insertMenu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(4);
    expect(deleteMenu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(6);
    expect(insertMenu.querySelector('path[fill="#107c41"]')).toBeTruthy();
    expect(deleteMenu.querySelector('path[fill="#c00000"]')).toBeTruthy();

    expect(menusCss).toMatch(
      /#menu-insert-cells,[\s\S]*?#menu-delete-cells\s*\{[\s\S]*?min-width: 166px;/,
    );
    expect(menusCss).toMatch(
      /#menu-insert-cells \.fc-tb__menu-item,[\s\S]*?#menu-delete-cells \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-insert-cells \.fc-tb__menu-icon,[\s\S]*?#menu-delete-cells \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Format Cells dropdown section chrome close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createFormatCellsMenu();
    const headings = Array.from(menu.children)
      .filter((child): child is HTMLElement => child instanceof HTMLElement)
      .filter((child) => child.classList.contains('fc-tb__menu-heading'))
      .map((heading) => heading.textContent);
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-cell-format]'));
    const submenuTriggers = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('[data-format-submenu]'),
    );
    const visibilityItems = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('#menu-format-cells-visibility [data-cell-format]'),
    );
    const tabColorMenu = menu.querySelector<HTMLElement>('#menu-format-cells-tabColor');
    const tabColorHeadings = Array.from(
      tabColorMenu?.querySelectorAll<HTMLElement>('.fc-tb__menu-heading') ?? [],
    ).map((heading) => heading.textContent);

    expect(menu.id).toBe('menu-format-cells');
    expect(headings).toEqual(['セルのサイズ', '表示設定', 'シートの整理', '保護']);
    expect(submenuTriggers.map((item) => item.dataset.formatSubmenu)).toEqual([
      'visibility',
      'tabColor',
    ]);
    expect(submenuTriggers.map((item) => item.textContent)).toEqual([
      '非表示/再表示',
      'シート見出しの色',
    ]);
    expect(menu.querySelector('#menu-format-cells-visibility')).toBeTruthy();
    expect(menu.querySelector('#menu-format-cells-tabColor')).toBeTruthy();
    expect(visibilityItems.map((item) => item.dataset.cellFormat)).toEqual([
      'hide-rows',
      'hide-cols',
      'hide-sheet',
      'show-rows',
      'show-cols',
      'unhide-sheet',
    ]);
    expect(visibilityItems.map((item) => item.textContent)).toEqual([
      '行を表示しない',
      '列を表示しない',
      'シートを表示しない',
      '行の再表示',
      '列の再表示',
      'シートの再表示...',
    ]);
    expect(menu.querySelector('[data-cell-format="rename-sheet"]')?.textContent).toBe(
      'シート名の変更',
    );
    expect(items.map((item) => item.dataset.cellFormat)).toContain('move-sheet-copy');
    expect(items.map((item) => item.dataset.cellFormat)).not.toContain('move-sheet-left');
    expect(items.map((item) => item.dataset.cellFormat)).not.toContain('move-sheet-right');
    expect(tabColorHeadings).toEqual(['テーマの色', '標準の色']);
    expect(tabColorMenu?.querySelector('[data-cell-format="tab-color-none"]')?.textContent).toBe(
      '色なし',
    );
    expect(tabColorMenu?.querySelector('[data-cell-format="tab-color-more"]')?.textContent).toBe(
      'その他の色…',
    );
    expect(tabColorMenu?.querySelectorAll('.fc-tb__color-swatch')).toHaveLength(14);
    expect(items.at(-1)?.dataset.cellFormat).toBe('dialog');
    expect(items.at(-1)?.textContent).toBe('セルの書式設定...');
    expect(items.map((item) => item.dataset.cellFormat)).toContain('lock-cell');
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg').length).toBeGreaterThan(
      10,
    );

    expect(menusCss).toMatch(/#menu-format-cells\s*\{[\s\S]*?min-width: 208px;/);
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__menu-heading\s*\{[\s\S]*?color: #808080;[\s\S]*?font-weight: 400;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__menu-item--checked \.fc-tb__menu-item__text::before\s*\{[\s\S]*?border-bottom: 2px solid #107c41;[\s\S]*?border-left: 2px solid #107c41;[\s\S]*?content: "";[\s\S]*?transform: rotate\(-45deg\);/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__submenu--format\s*\{[\s\S]*?min-width: 172px;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__submenu--format-tab-color\s*\{[\s\S]*?min-width: 206px;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__submenu--format-tab-color \.fc-tb__color-swatch-grid\s*\{[\s\S]*?grid-template-columns: repeat\(7, 18px\);/,
    );
  });

  it('keeps Merge Cells dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createMergeMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-merge-action]'));

    expect(menu.id).toBe('menu-merge');
    expect(items.map((item) => item.dataset.mergeAction)).toEqual([
      'mergeCenter',
      'mergeAcross',
      'mergeCells',
      'unmergeCells',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      'セルを結合して中央揃え',
      '横方向に結合',
      'セルの結合',
      'セル結合の解除',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(4);

    expect(menusCss).toMatch(/#menu-merge\s*\{[\s\S]*?min-width: 206px;/);
    expect(menusCss).toMatch(
      /#menu-merge \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-merge \.fc-tb__menu-icon,[\s\S]*?#menu-merge \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Wrap Text dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createWrapMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-wrap-action]'));

    expect(menu.id).toBe('menu-wrap');
    expect(items.map((item) => item.dataset.wrapAction)).toEqual(['wrapText', 'shrinkToFit']);
    expect(items.map((item) => item.textContent)).toEqual([
      '折り返して全体を表示',
      '縮小して全体を表示する',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon .fc-tb__menu-icon-svg')).toHaveLength(1);
    expect(menu.querySelectorAll('.fc-tb__menu-item__icon-spacer')).toHaveLength(1);

    expect(menusCss).toMatch(/#menu-wrap\s*\{[\s\S]*?min-width: 220px;/);
    expect(menusCss).toMatch(
      /#menu-wrap \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-wrap \.fc-tb__menu-icon,[\s\S]*?#menu-wrap \.fc-tb__menu-icon-svg,[\s\S]*?#menu-wrap \.fc-tb__menu-item__icon-spacer\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('adds stable modifier classes to Conditional Formatting icon-set choices', () => {
    const menu = createConditionalMenu('ja');
    const arrows = menu.querySelector<HTMLElement>('[data-cf-action="icons-arrows5"]');
    const traffic = menu.querySelector<HTMLElement>('[data-cf-action="icons-traffic3"]');
    const flags = menu.querySelector<HTMLElement>('[data-cf-action="icons-flags3"]');
    const symbols = menu.querySelector<HTMLElement>('[data-cf-action="icons-symbols3"]');

    expect(arrows?.classList.contains('fc-tb__cf-icon-choice--icons-arrows5')).toBe(true);
    expect(traffic?.classList.contains('fc-tb__cf-icon-choice--icons-traffic3')).toBe(true);
    expect(flags?.classList.contains('fc-tb__cf-icon-choice--icons-flags3')).toBe(true);
    expect(arrows?.querySelectorAll('span')).toHaveLength(5);
    expect(traffic?.querySelectorAll('span')).toHaveLength(3);
    expect(symbols?.querySelectorAll('span')).toHaveLength(3);
    expect(arrows?.textContent).toBe('');
    expect(symbols?.textContent).toBe('');

    const conditionalSource = readFileSync(
      join(root, 'src/toolbar/ribbon/menus/conditional.ts'),
      'utf8',
    );
    for (const glyph of ['▲', '↗', '▶', '↘', '▼', '★', '✓', '×', '⚑', '◔', '▮', '■']) {
      expect(conditionalSource).not.toContain(glyph);
    }
  });

  it('shows SVG previews for Borders dropdown footer and drawing tools', () => {
    const menu = createBordersMenu({
      ribbonText: {
        bottomBorder: 'Bottom Border',
        topBorder: 'Top Border',
        leftBorder: 'Left Border',
        rightBorder: 'Right Border',
        noBorder: 'No Border',
        allBorders: 'All Borders',
        outsideBorders: 'Outside Borders',
        thickOutsideBorders: 'Thick Outside Borders',
        doubleBottomBorder: 'Double Bottom Border',
        thickBottomBorder: 'Thick Bottom Border',
        topAndBottomBorder: 'Top and Bottom Border',
        topAndThickBottomBorder: 'Top and Thick Bottom Border',
        topAndDoubleBottomBorder: 'Top and Double Bottom Border',
        drawBordersHeading: 'Draw Borders',
        drawBorder: 'Draw Border',
        drawBorderGrid: 'Draw Border Grid',
        eraseBorder: 'Erase Border',
        lineColor: 'Line Color',
        lineStyle: 'Line Style',
        lineStyleNone: 'None',
        moreBorders: 'More Borders...',
        themeColors: 'Theme Colors',
        standardColors: 'Standard Colors',
        automatic: 'Automatic',
      } as Parameters<typeof createBordersMenu>[0]['ribbonText'],
      getBorderColor: () => '#000000',
      onPickColor: vi.fn(),
    });

    expect(menu.querySelector('[data-border-preset="format"] .fc-tb__border-preview')).toBeTruthy();
    expect(
      menu.querySelector('[data-border-draw="erase"] .fc-tb__border-preview--eraser'),
    ).toBeTruthy();
    expect(
      menu.querySelector('[data-border-submenu="lineColor"] .fc-tb__border-preview--line-color'),
    ).toBeTruthy();
    expect(
      menu.querySelector('[data-border-submenu="lineStyle"] .fc-tb__border-preview--line-style'),
    ).toBeTruthy();
  });

  it('keeps Borders dropdown close to Japanese Excel 365 desktop menu structure', () => {
    const menuCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const menu = createBordersMenu({
      ribbonText: {
        bottomBorder: '下罫線',
        topBorder: '上罫線',
        leftBorder: '左罫線',
        rightBorder: '右罫線',
        noBorder: '罫線なし',
        allBorders: '格子',
        outsideBorders: '外枠',
        thickOutsideBorders: '外枠太罫線',
        doubleBottomBorder: '下二重罫線',
        thickBottomBorder: '下太罫線',
        topAndBottomBorder: '上罫線 + 下罫線',
        topAndThickBottomBorder: '上罫線 + 下太罫線',
        topAndDoubleBottomBorder: '上罫線 + 下二重罫線',
        drawBordersHeading: '罫線の作成',
        drawBorder: '罫線の作成',
        drawBorderGrid: '罫線グリッドの作成',
        eraseBorder: '罫線の削除',
        lineColor: '線の色',
        lineStyle: '線のスタイル',
        lineStyleNone: 'なし',
        moreBorders: 'その他の罫線...',
        themeColors: 'テーマの色',
        standardColors: '標準の色',
        automatic: '自動',
      } as Parameters<typeof createBordersMenu>[0]['ribbonText'],
      getBorderColor: () => '#000000',
      onPickColor: vi.fn(),
    });

    expect(menu.classList.contains('fc-tb__menu--borders')).toBe(true);
    expect(menu.querySelectorAll('[role="separator"]')).toHaveLength(3);
    expect(menu.querySelector('.fc-tb__menu-heading')?.textContent).toBe('罫線の作成');
    expect(menu.querySelectorAll('[data-border-preset]')).toHaveLength(14);
    expect(menu.querySelectorAll('[data-border-draw]')).toHaveLength(3);
    expect(menu.querySelectorAll('[data-border-submenu]')).toHaveLength(2);
    expect(
      menu.querySelector('.fc-tb__submenu--line-style .fc-tb__submenu-item--line-style-none')
        ?.textContent,
    ).toBe('なし');
    expect(menu.querySelectorAll('.fc-tb__submenu--line-style .fc-tb__line-sample')).toHaveLength(
      13,
    );

    expect(menuCss).toMatch(/\.fc-tb__menu--borders\s*\{[\s\S]*?min-width: 186px;/);
    expect(menuCss).toMatch(
      /\.fc-tb__menu--borders \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 24px;[\s\S]*?padding: 2px 12px 2px 18px;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__menu--borders \.fc-tb__menu-item\[aria-expanded="true"\]\s*\{[\s\S]*?background: #107c41;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__menu--borders \.fc-tb__submenu--line-style\s*\{[\s\S]*?min-width: 101px;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__submenu--line-style \.fc-tb__line-sample\s*\{[\s\S]*?width: 74px;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__submenu--line-color \.fc-colorpalette__action--automatic\s*\{[\s\S]*?border-color: #107c41;/,
    );
  });
});
