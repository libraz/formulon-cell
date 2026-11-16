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
      '積み上げ縦棒',
      '100% 積み上げ縦棒',
      '## 3-D 縦棒',
      '[DIS]3-D 集合縦棒',
      '[DIS]3-D 積み上げ縦棒',
      '[DIS]3-D 100% 積み上げ縦棒',
      '[DIS]3-D 縦棒グラフ',
      '## 2-D 横棒',
      '集合横棒',
      '積み上げ横棒',
      '100% 積み上げ横棒',
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

  it('Recommended Charts offers its explanation until the selection holds data', () => {
    const menu = factory('mac.insert.recommendedChart');
    const none = menu?.querySelector<HTMLElement>(
      '[data-ribbon-command="mac.insert.recommendedChart.none"]',
    );
    expect(none?.textContent).toBe('選択したデータに対するおすすめグラフはありません。');
    expect(none?.getAttribute('aria-disabled')).toBe('true');
    expect(menu?.querySelector('[data-ribbon-command="mac.insert.chartColumn"]')).toBeNull();
  });

  it('words the German Recommended Charts menu with the app strings, not English', () => {
    const menu = createMacRibbonMenuFactory('de')('mac.insert.recommendedChart');
    const text = (id: string) =>
      menu?.querySelector<HTMLElement>(`[data-ribbon-command="${id}"]`)?.textContent;
    expect(text('mac.insert.recommendedChart.none')).toBe(
      'Für die von Ihnen ausgewählten Daten sind keine empfohlenen Diagramme vorhanden.',
    );
    expect(text('mac.insert.recommendedChart.apply')).toBe('Empfohlene Diagramme');
  });

  it('lists the Data Types and From Picture entries greyed out without host support', () => {
    const types = factory('mac.data.dataTypes');
    expect(rows(types as HTMLElement)).toEqual([
      '[DIS]株式 (English)',
      '[DIS]通貨 (English)',
      '[DIS]地理 (English)',
    ]);
    expect(rows(factory('mac.data.pictureFromData') as HTMLElement)).toEqual([
      '[DIS]ファイルからの画像...',
      '[DIS]クリップボードからの画像',
    ]);
  });
});
