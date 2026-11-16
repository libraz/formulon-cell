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

  it('applies the Mac row set and wording in every language', () => {
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
    expect(mac.slice(0, 2)).toEqual(['Ascending', 'Descending']);
    expect(mac.slice(2)).toEqual(shared.filter((text) => mac.slice(2).includes(text)));
  });

  it('words a German Mac menu with the app strings, not English', () => {
    const de = createHomeMenuFactories({
      ribbonLang: 'de',
      ribbonMenuText: dictionaries.de.ribbonMenu,
      ribbonText: toolbarText('de'),
      formatDialog: dictionaries.de.formatDialog,
      sheetTabs: dictionaries.de.sheetTabs,
      viewToolbar: dictionaries.de.viewToolbar,
    });
    expect(itemsOf(de.createSortMenu('sortFilterHome', MAC)).slice(0, 2)).toEqual([
      'Aufsteigend',
      'Absteigend',
    ]);
    expect(itemsOf(de.createInsertCellsMenu('insertRows', MAC))).toContain('Zellen einfügen...');
    expect(itemsOf(de.createFindSelectMenu('findHome', MAC))).toContain('Notizen');
  });

  it('words the German Mac Paste menu, its heading and its text rows with the app strings', () => {
    const menu = createPasteMenu(dictionaries.de, 'de', MAC);
    expect(menu.firstElementChild?.textContent).toBe('Einfügeoptionen:');
    const items = itemsOf(menu);
    expect(items).toContain('Nur den Text übernehmen');
    expect(items).toContain('Textimport-Assistenten verwenden...');
    expect(items.at(-1)).toBe('Inhalte einfügen...');
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

  it('Merge & Center carries Center Across Selection as a fifth row', () => {
    expect(itemsOf(factories.createMergeMenu('merge', MAC)).at(-1)).toBe('選択範囲内で中央');
    expect(itemsOf(factories.createMergeMenu('merge'))).toHaveLength(4);
  });

  it('Format menu has greyed captions, Default Width and a single Lock Cell row', () => {
    const menu = factories.createFormatCellsMenu('formatCellsHome', MAC);
    const captions = Array.from(
      menu.querySelectorAll<HTMLElement>(':scope > .fc-tb__menu-heading'),
      (el) => [(el.textContent ?? '').trim(), el.getAttribute('aria-disabled')],
    );
    expect(captions).toEqual([
      ['セルのサイズ', 'true'],
      ['表示設定', 'true'],
      ['シートの整理', 'true'],
      ['保護', 'true'],
    ]);
    const items = itemsOf(menu);
    expect(items).toContain('既定の幅...');
    expect(items).toContain('シートの保護...');
    expect(items.filter((label) => label.startsWith('セルのロック'))).toEqual(['セルのロック']);
    const desktop = itemsOf(factories.createFormatCellsMenu('formatCellsHome'));
    expect(desktop).not.toContain('既定の幅...');
    expect(desktop).toContain('セルのロック解除');
  });

  const sequence = (menu: HTMLElement): string[] =>
    Array.from(menu.children, (el) =>
      el.classList.contains('fc-tb__menu-sep') ? '---' : (el.textContent ?? '').trim(),
    );

  it('Fill has no separator on the Mac and Clear sets Remove Hyperlinks apart', () => {
    expect(sequence(factories.createFillMenu('fillHome', MAC))).not.toContain('---');
    expect(sequence(factories.createFillMenu('fillHome'))).toContain('---');
    const clear = sequence(factories.createClearMenu('clearHome', MAC));
    expect(clear.at(-2)).toBe('---');
    expect(clear.at(-1)).toBe('ハイパーリンクの削除');
  });

  it('Format leads its protection group with Protect Sheet on the Mac', () => {
    const rows = itemsOf(factories.createFormatCellsMenu('formatCellsHome', MAC));
    const protect = rows.indexOf('シートの保護...');
    expect(protect).toBeGreaterThan(-1);
    expect(rows[protect + 1]).toBe('セルのロック');
  });
});
