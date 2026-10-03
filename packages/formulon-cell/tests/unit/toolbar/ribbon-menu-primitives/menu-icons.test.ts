import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { EXCEL_RIBBON_ICON_PATHS } from '../../../../src/toolbar/excel-ribbon-icons.js';
import {
  MENU_ICON_GLYPHS,
  type MenuIconSlug,
  menuIconButton,
  setMenuIcon,
  VISUAL_TILE_GLYPHS,
  type VisualTileIconSlug,
  visualMenuTile,
} from '../../../../src/toolbar/ribbon/menus/general.js';
import { createTextOrientationMenu } from '../../../../src/toolbar/ribbon/menus/text-orientation.js';
import { readCssWithImports } from '../../../helpers/css.js';
import { collectStringLiteralArgs, collectVisualMenuTileIcons, root } from './fixtures.js';

const menuIconSlugs = Object.keys(MENU_ICON_GLYPHS) as MenuIconSlug[];
const visualTileSlugs = Object.keys(VISUAL_TILE_GLYPHS) as VisualTileIconSlug[];

const expectGlyph = (
  icon: Element | null | undefined,
  svgClass: string,
  glyph: string,
  slug: string,
): void => {
  const svg = icon?.querySelector(`.${svgClass}`);
  expect(svg, slug).toBeTruthy();
  expect(icon?.children.length, slug).toBe(1);
  expect(
    Array.from(svg?.querySelectorAll('path') ?? []).map((path) => path.getAttribute('d')),
    slug,
  ).toEqual(EXCEL_RIBBON_ICON_PATHS[glyph]?.map((segment) => segment.d));
};

describe('toolbar/ribbon menu primitives', () => {
  it('expands ribbon menu CSS imports in their declared order', () => {
    const menuCssFiles = [
      'menu-shells.css',
      'symbol-color-grids.css',
      'visual-galleries.css',
      'conditional-presets.css',
      'style-galleries.css',
      'conditional-menu.css',
      'ribbon-break.css',
      'menu-icons.css',
    ];
    const menuStylesDir = join(root, 'src/styles/toolbar/ribbon');
    const barrelPath = join(menuStylesDir, 'menus.css');
    const expected = menuCssFiles
      .map((file) => readCssWithImports(join(menuStylesDir, file)))
      .join('');

    expect(readCssWithImports(barrelPath)).toBe(expected);
    expect(readFileSync(barrelPath, 'utf8')).toBe(
      menuCssFiles.map((file) => `@import "./${file}";\n`).join(''),
    );
  });

  it('maps every menu icon slug to a registered icon-kit glyph', () => {
    expect(menuIconSlugs.length).toBeGreaterThan(150);
    for (const slug of menuIconSlugs) {
      const glyph = MENU_ICON_GLYPHS[slug];
      expect(EXCEL_RIBBON_ICON_PATHS[glyph]?.length ?? 0, `${slug} -> ${glyph}`).toBeGreaterThan(0);

      const button = menuIconButton(slug, 'auditAction', slug, slug);
      const icon = button.querySelector('.fc-tb__menu-icon');
      expect(icon?.classList.contains(`fc-tb__menu-icon--${slug}`), slug).toBe(true);
      expectGlyph(icon, 'fc-tb__menu-icon-svg', glyph, slug);
    }
  });

  it('keeps every literal menuIconButton slug in the glyph map', () => {
    const slugs = collectStringLiteralArgs('menuIconButton', 3);

    expect(slugs.length).toBeGreaterThan(100);
    for (const slug of slugs) expect(MENU_ICON_GLYPHS, slug).toHaveProperty([slug]);
  });

  it('gives calculation options and unfreeze their own glyphs', () => {
    expect({
      'calc-auto': MENU_ICON_GLYPHS['calc-auto'],
      'calc-auto-no-table': MENU_ICON_GLYPHS['calc-auto-no-table'],
      'calc-manual': MENU_ICON_GLYPHS['calc-manual'],
      'calc-now': MENU_ICON_GLYPHS['calc-now'],
      'calc-sheet': MENU_ICON_GLYPHS['calc-sheet'],
      'calc-iterative': MENU_ICON_GLYPHS['calc-iterative'],
      'freeze-off': MENU_ICON_GLYPHS['freeze-off'],
    }).toEqual({
      'calc-auto': 'calcAuto',
      'calc-auto-no-table': 'calcAutoNoTable',
      'calc-manual': 'calcManual',
      'calc-now': 'calcNow',
      'calc-sheet': 'calcSheet',
      'calc-iterative': 'calcIterative',
      'freeze-off': 'unfreeze',
    });
  });

  it('swaps both the slug class and the glyph when a menu icon is retargeted', () => {
    const button = menuIconButton('Freeze Panes', 'freeze', 'selection', 'freeze-panes');
    const icon = button.querySelector<HTMLElement>('.fc-tb__menu-icon');
    if (!icon) throw new Error('missing icon span');

    setMenuIcon(icon, 'freeze-off');
    expect(icon.classList.contains('fc-tb__menu-icon--freeze-panes')).toBe(false);
    expect(icon.classList.contains('fc-tb__menu-icon--freeze-off')).toBe(true);
    expectGlyph(icon, 'fc-tb__menu-icon-svg', 'unfreeze', 'freeze-off');

    setMenuIcon(icon, 'freeze-panes');
    expect(icon.classList.contains('fc-tb__menu-icon--freeze-off')).toBe(false);
    expectGlyph(icon, 'fc-tb__menu-icon-svg', 'freeze', 'freeze-panes');
  });

  it('draws no menu or tile icon artwork in CSS pseudo-elements', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).not.toMatch(/\.fc-tb__menu-icon--[a-z-]+::(before|after)/);
    expect(menusCss).not.toMatch(/\.fc-tb__visual-tile__icon--[a-z-]+::(before|after)/);
    expect(menusCss).toMatch(
      /\.fc-tb__menu-item__icon-spacer\s*\{[\s\S]*?flex: 0 0 18px;[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon\s*\{[\s\S]*?flex: 0 0 18px;[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('renders conditional-formatting symbol icons as vector marks, not text glyphs', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--check-green::before\s*\{[\s\S]*?border-bottom: 2px solid currentColor;[\s\S]*?border-left: 2px solid currentColor;[\s\S]*?transform: rotate\(-45deg\);/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--bang-yellow::before\s*\{[\s\S]*?width: 2px;[\s\S]*?height: 7px;[\s\S]*?background: currentColor;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--x-red::before,[\s\S]*?\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--x-red::after\s*\{[\s\S]*?width: 2px;[\s\S]*?height: 10px;[\s\S]*?background: currentColor;/,
    );
    const symbolCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__cf-icon--symbol::before'),
      menusCss.indexOf('.fc-tb__cf-icon--flag::before'),
    );
    expect(symbolCss).not.toContain('content: "✓"');
    expect(symbolCss).not.toContain('content: "!"');
    expect(symbolCss).not.toContain('content: "×"');
  });

  it('uses semantic SVGs for underline variant menu icons', () => {
    for (const iconSlug of ['underline-single', 'underline-double'] as const) {
      const button = menuIconButton('下線', 'underlineAction', 'single', iconSlug);
      const icon = button.querySelector('.fc-tb__menu-icon');

      expect(icon?.querySelector('.fc-tb__menu-icon-svg')).toBeTruthy();
      // The rule under the letterform is the underline itself, so it carries
      // the ink colour rather than an accent.
      expect(icon?.querySelector('path[fill="#1f1f1f"]')).toBeTruthy();
    }
  });

  it('maps every visual tile slug to a registered icon-kit glyph', () => {
    for (const slug of visualTileSlugs) {
      const button = visualMenuTile({ label: slug, attr: 'visualAction', value: slug, icon: slug });
      const icon = button.querySelector('.fc-tb__visual-tile__icon');

      expect(icon?.classList.contains(`fc-tb__visual-tile__icon--${slug}`), slug).toBe(true);
      expectGlyph(icon, 'fc-tb__visual-tile__icon-svg', VISUAL_TILE_GLYPHS[slug], slug);
    }
  });

  it('keeps every real visualMenuTile icon slug in the glyph map', () => {
    const slugs = collectVisualMenuTileIcons();

    expect(slugs.length).toBeGreaterThan(20);
    for (const slug of slugs) expect(VISUAL_TILE_GLYPHS, slug).toHaveProperty([slug]);
  });

  it('renders Text Orientation menu previews as colored semantic paths', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const menu = createTextOrientationMenu({
      orientationAngleCounterclockwise: '左回りに回転',
      orientationAngleClockwise: '右回りに回転',
      orientationVerticalText: '縦書き',
      orientationRotateTextUp: '上へ回転',
      orientationRotateTextDown: '下へ回転',
      orientationFormatAlignment: 'セルの配置の設定',
    } as Parameters<typeof createTextOrientationMenu>[0]);
    const previews = Array.from(
      menu.querySelectorAll<SVGSVGElement>('.fc-tb__text-orientation-preview'),
    );
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-text-orientation]'));

    expect(menu.id).toBe('menu-text-orientation');
    expect(items.map((item) => item.dataset.textOrientation)).toEqual([
      'ccw',
      'cw',
      'vertical',
      'up',
      'down',
      'format',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '左回りに回転',
      '右回りに回転',
      '縦書き',
      '上へ回転',
      '下へ回転',
      'セルの配置の設定',
    ]);
    expect(previews).toHaveLength(6);
    for (const preview of previews) {
      expect(preview.querySelector('text')).toBeNull();
      expect(preview.querySelectorAll('path').length).toBeGreaterThan(2);
      expect(preview.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    }
    expect(previews.some((preview) => preview.querySelector('path[stroke="#2f75b5"]'))).toBe(true);
    expect(menusCss).toMatch(/#menu-text-orientation\s*\{[\s\S]*?min-width: 168px;/);
    expect(menusCss).toMatch(
      /#menu-text-orientation \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-text-orientation \.fc-tb__text-orientation-preview\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });
});
