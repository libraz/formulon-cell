import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { en } from '../../../src/i18n/strings.js';
import { createSparklineStyleGallery } from '../../../src/toolbar/ribbon/sparkline-style-gallery.js';
import { createTableStyleGallery } from '../../../src/toolbar/ribbon/table-style-gallery.js';

const vars = (gallery: HTMLElement): string[] =>
  ['--fc-strip-chip-w', '--fc-strip-chip-h', '--fc-strip-step', '--fc-strip-visible'].map((name) =>
    gallery.style.getPropertyValue(name),
  );

describe('Quick Styles strip geometry on the Mac', () => {
  // Excel for Mac: seven chips in each strip; table 70 x 54 stepping 71, sparkline 80 x 56
  // stepping 78.
  it('sizes the table and sparkline strips as measured, without page arrows', () => {
    const table = createTableStyleGallery({
      lang: 'ja',
      getInst: () => null,
      arrow: document.createElement('button'),
      menu: null,
      mac: true,
    });
    expect(vars(table)).toEqual(['70px', '54px', '71px', '7']);
    expect(table.querySelector('.fc-tb__tablestyle-quick-prev')).toBeNull();
    const sparkline = createSparklineStyleGallery({
      lang: 'en',
      text: en.ribbon.sparkline,
      getInst: () => null,
      arrow: document.createElement('button'),
      menu: null,
      mac: true,
    });
    expect(vars(sparkline)).toEqual(['80px', '56px', '78px', '7']);
  });

  it('steps chips by the strip pitch, overlapping where it is below the chip width', () => {
    const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
    const css = readFileSync(join(root, 'src/styles/toolbar/ribbon/pivot-design.css'), 'utf8');
    expect(css).toMatch(
      /\.fc-tb__tablestyle-quick-chip \+ \.fc-tb__tablestyle-quick-chip,[^{]*\{\s*margin-left: calc\(var\(--fc-strip-step\) - var\(--fc-strip-chip-w\)\);/,
    );
    expect(css).toMatch(/flex: 0 1 calc\(var\(--fc-strip-step\) \* var\(--fc-strip-visible\)\);/);
  });

  it('shows each strip more arrow as a bare chevron', () => {
    const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
    const css = readFileSync(join(root, 'src/styles/toolbar/ribbon/pivot-design.css'), 'utf8');
    const rules = [...css.replace(/\/\*[\s\S]*?\*\//g, '').matchAll(/([^{}]*)\{([^}]*)\}/g)].map(
      (m) => ({
        selectors: (m[1] ?? '').split(/,\s*\n/).map((sel) => sel.trim()),
        body: m[2] ?? '',
      }),
    );
    const styled = (selector: string, declaration: string): boolean =>
      rules.some((r) => r.selectors.includes(selector) && r.body.includes(declaration));
    for (const gallery of [
      'fc-tb__pivotstyle-gallery',
      'fc-tb__tablestyle-quick-gallery',
      'fc-tb__sparkline-quick-gallery',
    ]) {
      const arrow = `.${gallery} > .fc-tb__rb`;
      expect(styled(arrow, 'width: 18px'), gallery).toBe(true);
      expect(styled(arrow, 'min-height: 0'), gallery).toBe(true);
      expect(styled(`${arrow} > :is(span, svg)`, 'display: none'), gallery).toBe(true);
    }
  });
});
