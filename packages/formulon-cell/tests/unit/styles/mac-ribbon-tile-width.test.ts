import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/platform-mac.css'), 'utf8').replace(
  /\/\*[\s\S]*?\*\//g,
  '',
);

/** Split a selector list on its top-level commas, leaving `:where(a, b)` whole. */
const splitList = (list: string): string[] => {
  const out: string[] = [];
  let depth = 0;
  let current = '';
  for (const ch of list) {
    if (ch === '(') depth++;
    if (ch === ')') depth--;
    if (ch === ',' && depth === 0) {
      out.push(current);
      current = '';
    } else current += ch;
  }
  out.push(current);
  return out.map((s) => s.replace(/\s+/g, ' ').trim());
};

/** Declarations of the rules whose selector list holds `selector`, later rules winning. */
const decls = (selector: string): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const [, selectors = '', body = ''] of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
    if (!splitList(selectors).includes(selector)) continue;
    for (const d of body.split(';')) {
      const [name, value] = d.split(/:(.+)/).map((x) => x.trim());
      if (name && value) out[name] = value;
    }
  }
  return out;
};

const GROUPS =
  '[data-fc-platform="mac"] > .fc-tb__ribbon-shell :where(.fc-tb__ribbon-group--tiles, .fc-tb__ribbon-group--mixed)';
const TILE = `${GROUPS} .fc-tb__rb--wide:not(.fc-tb__rb--stacked)`;

// Desktop AX frames at 1920 pt, en: tiles abut (Review New Comment 429.5+55 = Delete 484.5;
// Insert Text Box 1361+38 = Header & Footer 1399), plain tiles are at least 38 pt (Table, Link,
// Icons, Slicer, Text Box) and menu tiles at least 50 pt (Pictures, Shapes, Notes, WordArt).
// Multi-word labels wrap to two lines: New Comment 55, Recommended PivotTables 81.5.
describe('Mac ribbon full-height tiles', () => {
  it('sit side by side with no gap on every tab', () => {
    expect(decls(`${GROUPS} > .fc-tb__ribbon-tools`)['column-gap']).toBe('0');
  });

  it('hug their label from a 38 pt minimum on every tab', () => {
    expect(decls(TILE)).toMatchObject({
      width: 'auto',
      'min-width': '38px',
      'max-width': 'none',
      'padding-right': '1px',
      'padding-left': '1px',
    });
  });

  it('keep menu tiles at least 50 pt wide', () => {
    expect(decls(`${TILE}:where(:has(> .fc-tb__rb-split-chevron))`)['min-width']).toBe('50px');
  });

  it('lift the shared one-line label cap so the fit can wrap labels to two lines', () => {
    expect(
      decls(`${TILE} > span:not(.fc-tb__rb-icon):not(.fc-tb__rb-split-chevron)`),
    ).toMatchObject({ width: 'auto', 'min-width': 'min-content', 'max-width': 'none' });
  });

  it('no longer limits the hugging rule to the Home tab', () => {
    expect(css).not.toContain(':where([data-ribbon-panel="home"])');
  });
});
