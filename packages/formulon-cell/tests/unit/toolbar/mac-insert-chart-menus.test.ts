import { describe, expect, it } from 'vitest';
import { createMacRibbonMenuFactory } from '../../../src/toolbar/ribbon/mac/menus.js';

const rows = (menu: HTMLElement): string[] =>
  Array.from(
    menu.querySelectorAll<HTMLElement>('.fc-tb__menu-heading, [data-ribbon-command]'),
    (el) => {
      const text = (el.textContent ?? '').trim();
      if (el.classList.contains('fc-tb__menu-heading')) return `## ${text}`;
      return `${el.getAttribute('aria-disabled') === 'true' ? '[DIS]' : ''}${text}`;
    },
  );

// Excel for Mac (ja): Insert > 縦棒 / 円 galleries; chart kinds the host cannot draw stay greyed out.
describe('Mac Insert chart galleries', () => {
  const factory = createMacRibbonMenuFactory('ja');

  it('sections the column menu into 2-D / 3-D column and bar', () => {
    const menu = factory('mac.insert.chartColumn');
    if (!menu) throw new Error('missing menu');
    expect(rows(menu)).toEqual([
      '## 2-D 縦棒',
      '集合縦棒',
      '[DIS]積み上げ縦棒',
      '[DIS]100% 積み上げ縦棒',
      '## 3-D 縦棒',
      '[DIS]3-D 集合縦棒',
      '[DIS]3-D 積み上げ縦棒',
      '[DIS]3-D 100% 積み上げ縦棒',
      '[DIS]3-D 縦棒グラフ',
      '## 2-D 横棒',
      '集合横棒',
      '[DIS]積み上げ横棒',
      '[DIS]100% 積み上げ横棒',
      '## 3-D 横棒',
      '[DIS]3-D 集合横棒',
      '[DIS]3-D 積み上げ横棒',
      '[DIS]3-D 100% 積み上げ横棒',
    ]);
  });

  it('offers the doughnut chart the host draws', () => {
    const menu = factory('mac.insert.chartPie');
    expect(rows(menu as HTMLElement)).toContain('ドーナツ');
    expect(
      menu
        ?.querySelector('[data-ribbon-command="mac.insert.chartDoughnut"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('false');
  });
});
