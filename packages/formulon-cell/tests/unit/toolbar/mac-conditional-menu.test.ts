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

  it('lists the data bars and colour scales in the Excel set and order', () => {
    const swatches = (panel: string): string[] =>
      Array.from(
        mac.querySelectorAll<HTMLElement>(`[data-cf-panel="${panel}"] [data-cf-action]`),
        (el) => el.dataset.cfAction ?? '',
      );
    expect(swatches('dataBar')).toEqual([
      'data-blue',
      'data-green',
      'data-red',
      'data-orange',
      'data-teal',
      'data-purple',
      'data-solid-blue',
      'data-solid-green',
      'data-solid-red',
      'data-solid-orange',
      'data-solid-teal',
      'data-solid-purple',
      'new-rule',
    ]);
    expect(swatches('colorScale').slice(0, 12)).toEqual([
      'scale-gyr',
      'scale-ryg',
      'scale-gwr',
      'scale-rwg',
      'scale-bwr',
      'scale-rwb',
      'scale-wr',
      'scale-rw',
      'scale-gw',
      'scale-wg',
      'scale-gy',
      'scale-yg',
    ]);
    const first = mac.querySelector<HTMLElement>('[data-cf-action="scale-gyr"]');
    expect(first?.getAttribute('aria-label') ?? first?.title).toBe('緑、黄、赤のカラー スケール');
  });

  it('lists the twenty Excel icon sets with their wording', () => {
    const sets = Array.from(
      mac.querySelectorAll<HTMLElement>('[data-cf-panel="iconSet"] [data-cf-action^="icons-"]'),
      (el) => [el.dataset.cfAction, el.getAttribute('aria-label') ?? el.title],
    );
    expect(sets).toHaveLength(20);
    expect(sets[0]).toEqual(['icons-arrows3', '3 つの矢印 (色分け)']);
    expect(sets[2]).toEqual(['icons-triangles3', '3 種類の三角形']);
    expect(sets.at(-1)).toEqual(['icons-boxes5', '5 種類のボックス']);
  });
});
