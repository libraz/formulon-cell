import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { createHomeMenuFactories } from '../../../src/toolbar/ribbon/menus/home.js';
import { createPasteMenu } from '../../../src/toolbar/ribbon/menus/paste.js';
import { toolbarText } from '../../../src/toolbar/ribbon-model.js';

const factories = createHomeMenuFactories({
  ribbonLang: 'ja',
  ribbonMenuText: dictionaries.ja.ribbonMenu,
  ribbonText: toolbarText('ja'),
  formatDialog: dictionaries.ja.formatDialog,
  sheetTabs: dictionaries.ja.sheetTabs,
  viewToolbar: dictionaries.ja.viewToolbar,
});

const itemsOf = (menu: HTMLElement): string[] =>
  Array.from(menu.querySelectorAll<HTMLElement>('.fc-tb__menu-item'), (el) =>
    (el.textContent ?? '').trim(),
  );

const MAC = { mac: true } as const;

// Excel for Mac (ja) Home dropdown wording.
describe('Mac Home menu wording', () => {
  it('Sort & Filter keeps the six desktop rows', () => {
    const menu = factories.createSortMenu('sortFilterHome', MAC);
    expect(itemsOf(menu)).toEqual([
      '昇順',
      '降順',
      'ユーザー設定の並べ替え...',
      'フィルター',
      'クリア',
      '再適用',
    ]);
  });

  it('Find & Select names Go To Special and the selection pane as the desktop does', () => {
    const menu = factories.createFindSelectMenu('findHome', MAC);
    const items = itemsOf(menu);
    expect(items).toContain('選択オプション...');
    expect(items).toContain('メモ');
    expect(items).toContain('選択ウインドウ...');
    expect(items).not.toContain('コメントとメモ');
  });

  it('Insert, Delete, Format, Clear and Wrap use the desktop wording', () => {
    const insert = factories.createInsertCellsMenu('insertRows', MAC);
    expect(itemsOf(insert)).toContain('セルの挿入...');
    const del = factories.createDeleteCellsMenu('deleteRows', MAC);
    expect(itemsOf(del)).toContain('セルの削除...');
    const format = factories.createFormatCellsMenu('formatCellsHome', MAC);
    expect(itemsOf(format)).toContain('列の幅に合わせて自動調整');
    const clear = factories.createClearMenu('clearFormat', MAC);
    expect(itemsOf(clear)).toContain('コメントとメモをクリア');
    expect(itemsOf(clear)).not.toContain('条件付き書式のクリア');
    const wrap = factories.createWrapMenu('wrap', MAC);
    expect(itemsOf(wrap)[0]).toBe('折り返して全体を表示する');
  });

  it('applies the Mac row set in every language and keeps the shared wording outside Japanese', () => {
    const en = createHomeMenuFactories({
      ribbonLang: 'en',
      ribbonMenuText: dictionaries.en.ribbonMenu,
      ribbonText: toolbarText('en'),
      formatDialog: dictionaries.en.formatDialog,
      sheetTabs: dictionaries.en.sheetTabs,
      viewToolbar: dictionaries.en.viewToolbar,
    });
    const shared = itemsOf(en.createSortMenu('sortFilterHome'));
    const mac = itemsOf(en.createSortMenu('sortFilterHome', MAC));
    expect(mac).toHaveLength(6);
    expect(mac).toEqual(shared.filter((text) => mac.includes(text)));
    expect(mac[0]).toBe(shared[0]);
  });

  it('builds the full shared menu when no variant is asked for', () => {
    expect(itemsOf(factories.createSortMenu('sortFilterHome'))).toContain('詳細設定...');
  });

  it('Paste opens with a greyed options heading and Excel wording', () => {
    const menu = createPasteMenu(dictionaries.ja, 'ja', MAC);
    const first = menu.firstElementChild as HTMLElement;
    expect(first.textContent).toBe('貼り付けのオプション:');
    expect(first.getAttribute('aria-disabled')).toBe('true');
    const items = itemsOf(menu);
    expect(items[0]).toBe('ペースト');
    expect(items.at(-1)).toBe('形式を選択してペースト...');
    expect(itemsOf(createPasteMenu(dictionaries.ja, 'ja'))[0]).toBe('貼り付け');
  });
});
