import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(
  join(root, 'src/styles/core/app/overlays/quick-analysis-and-charts.css'),
  'utf8',
).replace(/\/\*[\s\S]*?\*\//g, '');

const ruleBody = (selector: string): string => {
  for (const m of css.matchAll(/([^{}]*)\{([^}]*)\}/g)) {
    const selectors = (m[1] ?? '').split(/,\s*\n/).map((sel) => sel.trim());
    if (selectors.includes(selector)) return m[2] ?? '';
  }
  throw new Error(`no rule for ${selector}`);
};

// A long tab label (French "Graphiques sparkline") must not run into the next tab: a tab keeps
// its label's width and the tab row scrolls instead.
describe('Quick Analysis tabs', () => {
  it('keep their label width inside the scrolling tab row', () => {
    expect(ruleBody('.fc-quick__tabs')).toMatch(/overflow-x: auto;/);
    const tab = ruleBody('.fc-quick__tab');
    expect(tab).toMatch(/flex: 0 0 auto;/);
    expect(tab).not.toMatch(/min-width: 0;/);
  });
});
