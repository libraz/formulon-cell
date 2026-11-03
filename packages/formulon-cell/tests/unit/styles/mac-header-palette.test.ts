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

// Excel for Mac, light appearance: strip 249-251,
// separators (233,233,231), label ink 94, strip border (188,187,186). The
// captures are Display P3; these greys read the same in sRGB.
describe('Mac light header palette', () => {
  it('matches the measured strip, separator, label and border colours', () => {
    const block = macLightPalette();
    expect(token(block, '--fc-header-bg')).toBe('#fafafa');
    expect(token(block, '--fc-header-rule')).toBe('#e9e9e7');
    expect(token(block, '--fc-header-fg')).toBe('#5e5e5e');
    expect(token(block, '--fc-rule-strong')).toBe('#bcbbba');
  });
});
