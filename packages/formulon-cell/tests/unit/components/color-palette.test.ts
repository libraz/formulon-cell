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
});
