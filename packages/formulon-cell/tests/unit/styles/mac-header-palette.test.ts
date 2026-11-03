import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

/** The Mac light palette block of platform-mac.css. */
const macLightPalette = (): string => {
  const css = readFileSync(join(root, 'src/styles/platform-mac.css'), 'utf8');
  const start = css.indexOf('@layer fc.theme');
  return css.slice(start, css.indexOf('\n}\n', start));
};

const token = (block: string, name: string): string | undefined =>
  block.match(new RegExp(`${name}:\\s*([^;]+);`))?.[1]?.trim();

// Excel for Mac, light appearance: strip 249-251, label ink 94, strip border
// (188,187,186); separators shade from 251 at the outer edge to 221 beside the border.
// The captures are Display P3; these greys read the same in sRGB.
describe('Mac light header palette', () => {
  it('matches the measured strip, separator, label and border colours', () => {
    const block = macLightPalette();
    expect(token(block, '--fc-header-bg')).toBe('#fafafa');
    // Unselected strips: 253 at the outer edge to 247 beside the border.
    expect(token(block, '--fc-header-strip-outer')).toBe('#fdfdfd');
    expect(token(block, '--fc-header-strip-inner')).toBe('#f7f7f7');
    expect(token(block, '--fc-header-rule-outer')).toBe('#fbfbfb');
    expect(token(block, '--fc-header-rule')).toBe('#dddcda');
    expect(token(block, '--fc-header-fg')).toBe('#5e5e5e');
    expect(token(block, '--fc-rule-strong')).toBe('#bcbbba');
  });
});

// Selected headers: flat 243 fill, green label
// ink 35,115,68 (sRGB), no accent line, label geometry unchanged (not bold).
describe('Mac light selected header', () => {
  it('fills 243 with a regular-weight green label and no accent bar', () => {
    const block = macLightPalette();
    expect(token(block, '--fc-bg-header')).toBe('#f3f3f3');
    expect(token(block, '--fc-header-fg-active')).toBe('#217346');
    expect(token(block, '--fc-header-active-bar')).toBe('transparent');
    expect(token(block, '--fc-header-fg-active-weight')).toBe('400');
  });
});
