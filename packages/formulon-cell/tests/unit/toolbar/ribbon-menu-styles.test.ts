import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

const stylesDir = join(__dirname, '../../../src/styles');
const rule = (file: string, selector: string): string => {
  const css = readFileSync(join(stylesDir, file), 'utf8');
  const start = css.indexOf(`${selector} {`);
  expect(start, `${selector} in ${file}`).toBeGreaterThanOrEqual(0);
  return css.slice(start, css.indexOf('}', start));
};

describe('ribbon menu styles', () => {
  it('scrolls visual gallery menus within the viewport', () => {
    const body = rule('toolbar/ribbon/visual-galleries.css', '.fc-tb__menu--visual');
    expect(body).toContain('overflow-y: auto');
    expect(body).toContain('max-height: calc(100vh - 16px)');
  });

  it('wraps and scrolls a folded group popup', () => {
    const body = rule(
      'platform-mac.css',
      '.fc-tb__ribbon-group--collapsed.fc-tb__ribbon-group--open\n    > .fc-tb__ribbon-tools',
    );
    expect(body).toContain('flex-wrap: wrap');
    expect(body).toContain('overflow: auto');
  });

  it('sizes Sheet Options checkbox columns to their labels on one line', () => {
    const grid = rule('platform-mac.css', '.fc-tb__ribbon-group--checks\n    .fc-tb__ribbon-tools');
    expect(grid).toContain('grid-auto-columns: max-content');
    const label = rule(
      'platform-mac.css',
      '.fc-tb__ribbon-group--checks\n    .fc-tb__rb\n    > span:not(.fc-tb__rb-icon):not(.fc-tb__rb-split-chevron)',
    );
    expect(label).toContain('white-space: nowrap');
  });
});
