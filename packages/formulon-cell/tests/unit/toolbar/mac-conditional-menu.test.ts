import { describe, expect, it } from 'vitest';
import { createConditionalMenu } from '../../../src/toolbar/ribbon/menus/conditional.js';

const panelItems = (menu: HTMLElement, panel: string): string[] =>
  Array.from(
    menu.querySelectorAll<HTMLButtonElement>(`[data-cf-panel="${panel}"] > .fc-tb__menu-item`),
    (el) => `${el.disabled ? '[DIS]' : ''}${(el.textContent ?? '').trim()}`,
  );

// Excel for Mac (ja): Home > 条件付き書式 submenus.
describe('Mac conditional formatting menu', () => {
  const mac = createConditionalMenu('ja', { mac: true });

  it('drops Unique Values from the highlight rules', () => {
    expect(panelItems(mac, 'highlight')).toEqual([
      '指定の値より大きい...',
      '指定の値より小さい...',
      '指定の範囲内...',
      '指定の値に等しい...',
      '文字列...',
      '日付...',
      '重複する値...',
      'その他のルール...',
    ]);
  });

  it('lists the top/bottom rules in the desktop order with an ellipsis', () => {
    expect(panelItems(mac, 'topBottom')).toEqual([
      '上位 10 項目...',
      '上位 10%...',
      '下位 10 項目...',
      '下位 10%...',
      '平均より上...',
      '平均より下...',
      'その他のルール...',
    ]);
  });

  it('shows the table and PivotTable clears greyed out', () => {
    expect(panelItems(mac, 'clear')).toEqual([
      '選択したセルからルールをクリア',
      'シート全体からルールをクリア',
      '[DIS]このテーブルからルールをクリア',
      '[DIS]このピボットテーブルからルールをクリア',
    ]);
  });

  it('keeps the shared menu when no variant is asked for', () => {
    const shared = createConditionalMenu('ja');
    expect(panelItems(shared, 'highlight')).toContain('一意の値...');
    expect(panelItems(shared, 'clear')).toHaveLength(2);
  });
});
