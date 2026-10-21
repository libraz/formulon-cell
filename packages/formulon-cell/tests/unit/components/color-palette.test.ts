import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, vi } from 'vitest';

import {
  createColorPalette,
  normalizeHex,
  PALETTE_COLUMNS,
  STANDARD_COLORS,
  syncColorPaletteTheme,
  THEME_COLOR_COLUMNS,
  themeColorColumns,
} from '../../../src/components/color-palette.js';

const swatches = (root: HTMLElement): HTMLButtonElement[] => [
  ...root.querySelectorAll<HTMLButtonElement>('.fc-colorpalette__swatch'),
];
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('color palette', () => {
  it('normalizes hex strings for stable equality', () => {
    expect(normalizeHex(' ABC ')).toBe('#aabbcc');
    expect(normalizeHex('#0F6CBD')).toBe('#0f6cbd');
    expect(normalizeHex('transparent')).toBe('#transparent');
  });

  it('renders theme and standard swatches with accessible labels', () => {
    const palette = createColorPalette({
      themeLabel: 'Theme Colors',
      standardLabel: 'Standard Colors',
      ariaLabel: 'Font color',
      value: '#156082',
      onPick: vi.fn(),
    });

    expect(palette.el.getAttribute('role')).toBe('group');
    expect(palette.el.getAttribute('aria-label')).toBe('Font color');
    expect(swatches(palette.el)).toHaveLength(
      THEME_COLOR_COLUMNS.length * 6 + STANDARD_COLORS.length,
    );
    expect(
      palette.el
        .querySelector<HTMLButtonElement>('[data-color="#156082"]')
        ?.getAttribute('aria-pressed'),
    ).toBe('true');
    expect(
      palette.el
        .querySelector<HTMLButtonElement>('[data-color="#156082"]')
        ?.getAttribute('aria-label'),
    ).toBe('Accent 1 (#156082)');
    expect(
      palette.el
        .querySelector<HTMLButtonElement>('[data-color="#c0e6f5"]')
        ?.getAttribute('aria-label'),
    ).toBe('Accent 1, Lighter 80% (#c0e6f5)');
  });

  it('picks swatches, updates selected state, and normalizes setValue input', () => {
    const onPick = vi.fn();
    const palette = createColorPalette({
      themeLabel: 'Theme Colors',
      standardLabel: 'Standard Colors',
      value: '#000000',
      onPick,
    });
    const red = palette.el.querySelector<HTMLButtonElement>('[data-color="#ff0000"]');

    red?.click();

    expect(onPick).toHaveBeenCalledWith('#ff0000');
    expect(red?.getAttribute('aria-pressed')).toBe('true');

    palette.setValue('4EA72E');

    expect(
      palette.el
        .querySelector<HTMLButtonElement>('[data-color="#4ea72e"]')
        ?.getAttribute('aria-pressed'),
    ).toBe('true');
    expect(red?.getAttribute('aria-pressed')).toBe('false');
  });

  it('uses roving tabindex and keyboard navigation across the 10-column grid', () => {
    const palette = createColorPalette({
      themeLabel: 'Theme Colors',
      standardLabel: 'Standard Colors',
      value: '#ffffff',
      onPick: vi.fn(),
    });
    const buttons = swatches(palette.el);
    document.body.appendChild(palette.el);

    try {
      palette.focus();
      expect(document.activeElement).toBe(buttons[0]);
      expect(buttons[0]?.tabIndex).toBe(0);
      expect(buttons[1]?.tabIndex).toBe(-1);

      buttons[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));

      expect(document.activeElement).toBe(buttons[PALETTE_COLUMNS]);
      expect(buttons[PALETTE_COLUMNS]?.tabIndex).toBe(0);

      buttons[PALETTE_COLUMNS]?.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'End', ctrlKey: true, bubbles: true }),
      );

      expect(document.activeElement).toBe(buttons.at(-1));
    } finally {
      palette.el.remove();
    }
  });

  it('exposes Automatic and More Colors actions without changing swatch roving state', () => {
    const onPick = vi.fn();
    const onMoreColors = vi.fn();
    const palette = createColorPalette({
      themeLabel: 'Theme Colors',
      standardLabel: 'Standard Colors',
      value: '#ffffff',
      automatic: { label: 'Automatic', color: '#000000' },
      moreColorsLabel: 'More Colors...',
      onPick,
      onMoreColors,
    });

    palette.el.querySelector<HTMLButtonElement>('.fc-colorpalette__action--automatic')?.click();
    palette.el.querySelector<HTMLButtonElement>('.fc-colorpalette__action--more')?.click();

    expect(onPick).toHaveBeenCalledWith('#000000');
    expect(onMoreColors).toHaveBeenCalledTimes(1);
    expect(swatches(palette.el).filter((button) => button.tabIndex === 0)).toHaveLength(1);
  });

  it('can render the Excel fill-color high contrast row without making it focusable', () => {
    const palette = createColorPalette({
      themeLabel: 'テーマの色',
      standardLabel: '標準の色',
      value: '#ffffff',
      highContrastOnlyLabel: 'ハイ コントラストのみ',
      onPick: vi.fn(),
    });

    const contrast = palette.el.querySelector<HTMLLabelElement>('.fc-colorpalette__contrast');
    const input = contrast?.querySelector<HTMLInputElement>('input[type="checkbox"]');

    expect(contrast?.textContent).toContain('ハイ コントラストのみ');
    expect(input?.disabled).toBe(true);
    expect(input?.tabIndex).toBe(-1);
    expect(swatches(palette.el)).toHaveLength(
      THEME_COLOR_COLUMNS.length * 6 + STANDARD_COLORS.length,
    );
  });

  it('keeps the shared color palette close to Japanese Excel 365 desktop geometry', () => {
    const paletteCss = readFileSync(join(root, 'src/styles/core/app/color-palette.css'), 'utf8');

    expect(paletteCss).toMatch(/\.fc-colorpalette\s*\{[\s\S]*?padding: 7px 10px 8px;/);
    expect(paletteCss).toMatch(
      /\.fc-colorpalette__contrast\s*\{[\s\S]*?min-height: 20px;[\s\S]*?white-space: nowrap;/,
    );
    expect(paletteCss).toMatch(/\.fc-colorpalette__heading\s*\{[\s\S]*?font-size: 12px;/);
    expect(paletteCss).toMatch(
      /\.fc-colorpalette__grid\s*\{[\s\S]*?grid-template-columns: repeat\(10, 18px\);/,
    );
    expect(paletteCss).toMatch(
      /\.fc-colorpalette__swatch\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
    expect(paletteCss).toMatch(
      /\.fc-colorpalette__swatch:focus-visible\s*\{[\s\S]*?outline: 2px solid var\(--fc-accent\);/,
    );
    expect(paletteCss).toMatch(
      /\.fc-colorpalette__action--automatic\s*\{[\s\S]*?justify-content: center;[\s\S]*?border-color: var\(--fc-accent\);/,
    );
    expect(paletteCss).toMatch(
      /\.fc-colorpalette__action:hover,[\s\S]*?\.fc-colorpalette__action:focus-visible\s*\{[\s\S]*?background: var\(--fc-bg-hover, color-mix\(in srgb, CanvasText 8%, transparent\)\);/,
    );
    expect(paletteCss).toMatch(
      /\.fc-colorpalette__wheel\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
    expect(paletteCss).toMatch(/\.fc-colorpalette__action--more\s*\{[\s\S]*?margin: 6px -4px 0;/);
    expect(paletteCss).not.toContain('background: var(--fc-accent-soft');
  });

  it('tints the theme the way the spreadsheet paints each variant', () => {
    const [background1, text1, background2, text2, accent1] = themeColorColumns(undefined);
    // Cell colours the spreadsheet reports for the default theme at these tints.
    expect(accent1?.variants).toEqual(['#c0e6f5', '#83cceb', '#44b3e1', '#104861', '#0b3040']);
    expect(background1?.variants.slice(0, 1)).toEqual(['#f2f2f2']);
    expect(background1?.variants.at(-1)).toBe('#808080');
    expect(text1?.variants.slice(0, 2)).toEqual(['#808080', '#595959']);
    expect(background2?.variants.slice(0, 3)).toEqual(['#d0d0d0', '#adadad', '#747474']);
    expect(background2?.variants.at(-1)).toBe('#161616');
    expect([text2?.variants[0], text2?.variants[2]]).toEqual(['#dae9f8', '#4d93d9']);
  });

  it('shows a changed workbook theme in a palette already built', () => {
    const onPick = vi.fn();
    const palette = createColorPalette({
      themeLabel: 'Theme Colors',
      standardLabel: 'Standard Colors',
      onPick,
    });
    const accent1 = swatches(palette.el)[4];
    if (accent1) accent1.dataset.shapeFill = accent1.dataset.color;
    const gray = [
      0xff000000, 0xffffffff, 0xff262626, 0xfff2f2f2, 0xff7f7f7f, 0xffa6a6a6, 0xff595959,
      0xffbfbfbf, 0xffd9d9d9, 0xff404040,
    ];

    syncColorPaletteTheme(palette.el, themeColorColumns(gray));

    expect(accent1?.dataset.color).toBe('#7f7f7f');
    expect(accent1?.dataset.shapeFill).toBe('#7f7f7f');
    expect(accent1?.getAttribute('aria-label')).toBe('Accent 1 (#7f7f7f)');
    accent1?.click();
    expect(onPick).toHaveBeenCalledWith('#7f7f7f');
  });
});
