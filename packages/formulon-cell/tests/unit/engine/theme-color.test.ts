import { describe, expect, it } from 'vitest';
import { tintCssColor } from '../../../src/engine/theme-color.js';

const hex = (rgb: readonly [number, number, number]): string =>
  `#${rgb.map((v) => v.toString(16).padStart(2, '0')).join('')}`;

// Tinted theme colours the desktop app reports for its built-in table styles
// under the Office 2007 theme, with the tints as the file stores them.
const DESKTOP: readonly [base: string, tint: number, rgb: [number, number, number]][] = [
  ['#9bbb59', 0.7999816888943144, [235, 241, 222]],
  ['#9bbb59', 0.5999938962981048, [216, 228, 188]],
  ['#000000', 0.1499984740745262, [38, 38, 38]],
  ['#000000', 0.249977111117893, [64, 64, 64]],
  ['#4f81bd', 0.7999816888943144, [220, 230, 241]],
  ['#4f81bd', -0.499984740745262, [36, 64, 98]],
  ['#ffffff', -0.1499984740745262, [217, 217, 217]],
  ['#ffffff', -0.3499862666707358, [166, 166, 166]],
];

describe('theme colour tints', () => {
  it.each(DESKTOP)('tints %s by %d as the desktop app does', (base, tint, rgb) => {
    expect(tintCssColor(base, tint)).toBe(hex(rgb));
  });

  it('leaves a colour untouched at tint 0', () => {
    expect(tintCssColor('#4f81bd', 0)).toBe('#4f81bd');
  });
});
