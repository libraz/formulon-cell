import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { createPivotMenuFactories } from '../../../src/toolbar/ribbon/menus/pivot.js';

const menu = (commandId: string): HTMLDivElement =>
  createPivotMenuFactories({ ribbonMenuText: dictionaries.ja.ribbonMenu }).createPivotLayoutMenu(
    commandId,
  );

/** Visible rows: label, '---' for a separator, ' (dis)' when disabled. */
const rows = (el: HTMLElement): string[] =>
  Array.from(el.children).flatMap((child) => {
    if (child.classList.contains('fc-tb__menu-sep')) return ['---'];
    if (!(child instanceof HTMLButtonElement)) return [];
    const disabled = child.getAttribute('aria-disabled') === 'true' ? ' (dis)' : '';
    return [`${child.textContent?.trim()}${disabled}`];
  });

const values = (el: HTMLElement): string[] =>
  Array.from(el.querySelectorAll<HTMLButtonElement>('[data-pivot-layout]')).map(
    (b) => b.dataset.pivotLayout ?? '',
  );

// Mac Excel ja, Design and Analyze tab menus as measured.
describe('PivotTable Design and Analyze menus', () => {
  it('lists the subtotal positions, with Do Not Show and filtered items unavailable', () => {
    const el = menu('pivotSubtotals');
    expect(rows(el)).toEqual([
      '小計を表示しない (dis)',
      'すべての小計をグループの末尾に表示する',
      'すべての小計をグループの先頭に表示する',
      '---',
      'フィルターされたアイテムを合計に含める (dis)',
    ]);
    expect(values(el)).toEqual([
      'subtotals-none',
      'subtotals-bottom',
      'subtotals-top',
      'subtotals-filtered',
    ]);
  });

  it('lists the four grand total states', () => {
    expect(rows(menu('pivotGrandTotals'))).toEqual([
      '行と列の集計を行わない',
      '行と列の集計を行う',
      '行のみ集計を行う',
      '列のみ集計を行う',
    ]);
  });

  it('lists the three layouts and the label repeats, which are unavailable', () => {
    expect(rows(menu('pivotReportLayout'))).toEqual([
      'コンパクト形式で表示',
      'アウトライン形式で表示',
      '表形式で表示',
      '---',
      'すべてのアイテムのラベルを繰り返す (dis)',
      'アイテムのラベルを繰り返さない (dis)',
    ]);
  });

  it('lists blank lines, both unavailable', () => {
    expect(rows(menu('pivotBlankRows'))).toEqual([
      '各アイテムの後ろに空白行を挿入する (dis)',
      '各アイテムの後ろの空白行を削除する (dis)',
    ]);
  });

  it('lists Clear, Select and Fields, Items & Sets', () => {
    expect(rows(menu('pivotClear'))).toEqual(['すべてクリア', 'フィルターのクリア']);
    expect(rows(menu('pivotSelect'))).toEqual([
      'ラベルと値 (dis)',
      '値 (dis)',
      'ラベル (dis)',
      'ピボットテーブル全体',
      '---',
      '選択範囲を有効にする (dis)',
    ]);
    expect(
      menu('pivotSelect')
        .querySelector('[data-pivot-layout="select-enable"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
    expect(rows(menu('pivotFieldsItemsSets'))).toEqual([
      '集計フィールド... (dis)',
      '集計アイテム... (dis)',
      '解決の順序... (dis)',
      '数式の一覧表示 (dis)',
    ]);
  });

  it('gives every unavailable entry its reason', () => {
    const disabled = Array.from(
      menu('pivotReportLayout').querySelectorAll<HTMLButtonElement>('[aria-disabled="true"]'),
    );
    expect(
      disabled.every((b) => b.dataset.disabledReason === 'この操作はまだ使用できません。'),
    ).toBe(true);
  });
});
