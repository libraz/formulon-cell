import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, vi } from 'vitest';

import {
  createColorPalette,
  normalizeHex,
  PALETTE_COLUMNS,
  STANDARD_COLORS,
  THEME_COLOR_COLUMNS,
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
      value: '#4472c4',
      onPick: vi.fn(),
    });

    expect(palette.el.getAttribute('role')).toBe('group');
    expect(palette.el.getAttribute('aria-label')).toBe('Font color');
    expect(swatches(palette.el)).toHaveLength(
      THEME_COLOR_COLUMNS.length * 6 + STANDARD_COLORS.length,
    );
    expect(
      palette.el
        .querySelector<HTMLButtonElement>('[data-color="#4472c4"]')
        ?.getAttribute('aria-pressed'),
    ).toBe('true');
    expect(
      palette.el
        .querySelector<HTMLButtonElement>('[data-color="#4472c4"]')
        ?.getAttribute('aria-label'),
    ).toBe('Blue, Accent 1 (#4472c4)');
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

    palette.setValue('70AD47');

    expect(
      palette.el
        .querySelector<HTMLButtonElement>('[data-color="#70ad47"]')
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
});
