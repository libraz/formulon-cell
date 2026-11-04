import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { builtinTableStyleElements } from '../../../src/engine/builtin-table-styles.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import { createStylesMenuFactories } from '../../../src/toolbar/ribbon/menus/styles.js';

/** Swatch labels of the Mac Format as Table gallery, ja, in gallery order. */
const JA_SWATCHES = [
  '白, テーブル スタイル (淡色) 1',
  'アイスブルー, テーブル スタイル (淡色) 2',
  'ローズ, テーブル スタイル (淡色) 3',
  '薄い緑, テーブル スタイル (淡色) 4',
  'ラベンダー, テーブル スタイル (淡色) 5',
  'スカイブルー, テーブル スタイル (淡色) 6',
  '薄いオレンジ, テーブル スタイル (淡色) 7',
  '白, テーブル スタイル (淡色) 8',
  '青, テーブル スタイル (淡色) 9',
  '赤, テーブル スタイル (淡色) 10',
  'オリーブ, テーブル スタイル (淡色) 11',
  '紫, テーブル スタイル (淡色) 12',
  'アクア, テーブル スタイル (淡色) 13',
  'オレンジ, テーブル スタイル (淡色) 14',
  '白, テーブル スタイル (淡色) 15',
  'アイスブルー, テーブル スタイル (淡色) 16',
  'ローズ, テーブル スタイル (淡色) 17',
  '薄い緑, テーブル スタイル (淡色) 18',
  'ラベンダー, テーブル スタイル (淡色) 19',
  'スカイブルー, テーブル スタイル (淡色) 20',
  '薄いオレンジ, テーブル スタイル (淡色) 21',
  '白, テーブル スタイル (中間) 1',
  '青, テーブル スタイル (中間) 2',
  '赤, テーブル スタイル (中間) 3',
  'オリーブ, テーブル スタイル (中間) 4',
  '紫, テーブル スタイル (中間) 5',
  'アクア, テーブル スタイル (中間) 6',
  'オレンジ, テーブル スタイル (中間) 7',
  '薄い灰色, テーブル スタイル (中間) 8',
  '青, テーブル スタイル (中間) 9',
  '赤, テーブル スタイル (中間) 10',
  'オリーブ, テーブル スタイル (中間) 11',
  '紫, テーブル スタイル (中間) 12',
  'アクア, テーブル スタイル (中間) 13',
  'オレンジ, テーブル スタイル (中間) 14',
  '白, テーブル スタイル (中間) 15',
  '青, テーブル スタイル (中間) 16',
  '赤, テーブル スタイル (中間) 17',
  'オリーブ, テーブル スタイル (中間) 18',
  '紫, テーブル スタイル (中間) 19',
  'アクア, テーブル スタイル (中間) 20',
  'オレンジ, テーブル スタイル (中間) 21',
  '薄い灰色, テーブル スタイル (中間) 22',
  'アイスブルー, テーブル スタイル (中間) 23',
  'ローズ, テーブル スタイル (中間) 24',
  '薄い緑, テーブル スタイル (中間) 25',
  'ラベンダー, テーブル スタイル (中間) 26',
  'スカイブルー, テーブル スタイル (中間) 27',
  '薄いオレンジ, テーブル スタイル (中間) 28',
  '濃い灰色, テーブル スタイル (濃色) 1',
  '青, テーブル スタイル (濃色) 2',
  '赤, テーブル スタイル (濃色) 3',
  'オリーブ, テーブル スタイル (濃色) 4',
  '濃い紫, テーブル スタイル (濃色) 5',
  'アクア, テーブル スタイル (濃色) 6',
  'オレンジ, テーブル スタイル (濃色) 7',
  '薄い灰色, テーブル スタイル (濃色) 8',
  '赤, テーブル スタイル (濃色) 9',
  '紫, テーブル スタイル (濃色) 10',
  'オレンジ, テーブル スタイル (濃色) 11',
];

const gallery = (lang: 'ja' | 'en'): HTMLDivElement =>
  createStylesMenuFactories({
    ribbonLang: lang,
    ribbonMenuText: dictionaries[lang].ribbonMenu,
    ribbonText: dictionaries[lang].ribbon,
  }).createTableStyleMenu('formatTableHome');

const swatches = (menu: HTMLElement): HTMLButtonElement[] =>
  Array.from(menu.querySelectorAll<HTMLButtonElement>('.fc-tb__tablestyle-swatch'));

const names = (family: string, count: number): string[] =>
  Array.from({ length: count }, (_, i) => `TableStyle${family}${i + 1}`);

describe('Format as Table gallery', () => {
  it('lists the 60 built-in styles by name, Light 1-21, Medium 1-28, Dark 1-11', () => {
    const menu = gallery('ja');
    const grids = Array.from(menu.querySelectorAll<HTMLElement>('.fc-tb__tablestyle-grid'));
    expect(grids.map((grid) => grid.querySelectorAll('.fc-tb__tablestyle-swatch').length)).toEqual([
      21, 28, 11,
    ]);
    expect(swatches(menu).map((b) => b.dataset.tableStyle)).toEqual([
      ...names('Light', 21),
      ...names('Medium', 28),
      ...names('Dark', 11),
    ]);
  });

  it('names each ja swatch with its colour and number, as the Mac app does', () => {
    const menu = gallery('ja');
    expect(swatches(menu).map((b) => b.getAttribute('aria-label'))).toEqual(JA_SWATCHES);
    expect(swatches(menu).map((b) => b.title)).toEqual(JA_SWATCHES);
    expect(menu.textContent).not.toContain('なし');
    expect(
      Array.from(menu.querySelectorAll('.fc-tb__tablestyle-footer')).map((b) => b.textContent),
    ).toEqual(['新しい表スタイル...', '新しいピボットテーブル スタイル...']);
  });

  it('names swatches by family and number only where colour names are unmeasured', () => {
    const labels = swatches(gallery('en')).map((b) => b.getAttribute('aria-label'));
    expect(labels.slice(0, 2)).toEqual(['Table Style (Light) 1', 'Table Style (Light) 2']);
    expect(labels[21]).toBe('Table Style (Medium) 1');
    expect(labels.at(-1)).toBe('Table Style (Dark) 11');
  });

  it('leaves the table style options to the table: built-in swatches carry no variant', () => {
    for (const b of swatches(gallery('ja'))) expect(b.dataset.tableVariant).toBeUndefined();
  });

  it('lays Mac swatches out on the measured 71 x 56 pt pitch of 73 x 58 pt frames', () => {
    const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
    const css = readFileSync(join(root, 'src/styles/toolbar/ribbon/style-galleries.css'), 'utf8');
    const rule = (selector: string): string =>
      [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)].find((m) =>
        (m[1] ?? '')
          .replace(/\/\*[\s\S]*?\*\//g, '')
          .split(',')
          .some((x) => x.trim() === selector),
      )?.[2] ?? '';
    const grid = rule('[data-fc-platform="mac"] .fc-tb__tablestyle-grid');
    expect(grid).toContain('grid-template-columns: repeat(7, 71px)');
    expect(grid).toContain('grid-auto-rows: 56px');
    const swatch = rule('[data-fc-platform="mac"] .fc-tb__tablestyle-swatch');
    expect(swatch).toContain('width: 73px');
    expect(swatch).toContain('height: 58px');
    expect(swatch).toContain('margin: -1px');
  });

  it('draws a built-in swatch from its style definition', () => {
    const medium2 = gallery('ja').querySelector<HTMLElement>(
      '[data-table-style="TableStyleMedium2"]',
    );
    const cells = Array.from(
      medium2?.querySelectorAll<HTMLElement>('.fc-tb__tablestyle-swatch__cell') ?? [],
    );
    expect(cells).toHaveLength(25);
    // Header row as the definition paints it under the default theme.
    const header = builtinTableStyleElements('TableStyleMedium2')?.headerRow;
    expect(cells[0]?.style.background).toBe(header?.fill);
    expect(cells[0]?.style.color).toBe(header?.color);
    // The first and third body rows carry the first stripe; the others are plain.
    expect(cells[5]?.style.background).not.toBe('');
    expect(cells[10]?.style.background).toBe('');
    // Light 9 rules every body row through its two stripes.
    const light9 = gallery('ja').querySelector<HTMLElement>(
      '[data-table-style="TableStyleLight9"]',
    );
    const rows = Array.from(
      light9?.querySelectorAll<HTMLElement>('.fc-tb__tablestyle-swatch__cell') ?? [],
    ).filter((_, i) => i % 5 === 0 && i > 0);
    expect(rows.map((cell) => cell.style.borderTop !== '')).toEqual([true, true, true, true]);
  });
});
