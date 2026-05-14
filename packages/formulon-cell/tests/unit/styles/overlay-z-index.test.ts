import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const styleRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../src/styles/core');

const readStyle = (file: string): string => readFileSync(resolve(styleRoot, file), 'utf8');

const baseCss = readStyle('base.css');

// Recursively concat every .css file under `core/app/` — the CSS refactor
// splits overlay rules across an evolving file tree (root files, `app/dialogs/`,
// `app/overlays/`, `app/panels/`, `app/format-dialog/`, …). The test cares about
// z-index declarations and selector blocks, not file boundaries.
const walkCss = (dir: string): string[] => {
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    const full = resolve(dir, name);
    if (statSync(full).isDirectory()) out.push(...walkCss(full));
    else if (name.endsWith('.css')) out.push(readFileSync(full, 'utf8'));
  }
  return out;
};
const overlayCss = walkCss(resolve(styleRoot, 'app')).join('\n');

const expectedTierValues = {
  base: 2_147_483_000,
  grid: 2_147_483_010,
  dialog: 2_147_483_020,
  tooltip: 2_147_483_030,
  popover: 2_147_483_040,
  callout: 2_147_483_050,
  menu: 2_147_483_060,
  error: 2_147_483_070,
} as const;

type ZTier = Exclude<keyof typeof expectedTierValues, 'base'>;

const expectedStackOrder: ZTier[] = [
  'grid',
  'dialog',
  'tooltip',
  'popover',
  'callout',
  'menu',
  'error',
];

const overlayExpectations: [selector: string, tier: ZTier][] = [
  ['.fc-quick', 'grid'],
  ['.fc-charts', 'grid'],
  ['.fc-objects', 'grid'],
  ['.fc-find', 'grid'],
  ['.fc-slicer', 'grid'],
  ['.fc-fmtdlg', 'dialog'],
  ['.fc-iterdlg', 'dialog'],
  ['.fc-extlinkdlg', 'dialog'],
  ['.fc-cfrulesdlg', 'dialog'],
  ['.fc-stylegallery', 'dialog'],
  ['.fc-hover-tip', 'tooltip'],
  ['.fc-filter-dropdown', 'popover'],
  ['.fc-autocomplete', 'popover'],
  ['.fc-arghelper', 'popover'],
  ['.fc-validation-list', 'popover'],
  ['.fc-statusbar__chooser', 'popover'],
  ['.fc-sheetmenu', 'popover'],
  ['.fc-cmtnote', 'callout'],
  ['.fc-ctxmenu', 'menu'],
  ['.fc-errmenu', 'error'],
];

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const cssBlockFor = (css: string, selector: string): string => {
  // Find every `.selector { ... }` block (some selectors appear multiple
  // times — once in the main rule, again in a responsive override, etc.).
  // The lookahead prevents `.fc-fmtdlg` from matching `.fc-fmtdlg__panel`.
  const pattern = new RegExp(
    `${escapeRegExp(selector)}(?![\\w-])\\s*\\{([\\s\\S]*?)\\n\\s*\\}`,
    'gm',
  );
  const blocks: string[] = [];
  for (const m of css.matchAll(pattern)) blocks.push(m[1] ?? '');
  if (blocks.length === 0) throw new Error(`Missing CSS block for ${selector}`);
  // Prefer the block that declares a z-index — that's the "main" overlay rule
  // we're trying to inspect, not a responsive override.
  const withZ = blocks.find((b) => /z-index:/.test(b));
  return withZ ?? blocks[0] ?? '';
};

const zIndexVarPattern = /z-index:\s*var\(--fc-z-([a-z]+),\s*([0-9]+)\);/g;

describe('overlay z-index tiers', () => {
  it('defines the global overlay stack in the intended bottom-to-top order', () => {
    expect(baseCss).toMatch(/:where\(html\)\s*\{/);
    expect(baseCss).not.toMatch(/\.fc-host\s*\{[\s\S]*--fc-z-base:/);
    expect(baseCss).toContain(`--fc-z-base: ${expectedTierValues.base};`);

    for (const tier of expectedStackOrder) {
      const offset = expectedTierValues[tier] - expectedTierValues.base;
      expect(baseCss).toContain(`--fc-z-${tier}: calc(var(--fc-z-base) + ${offset});`);
    }

    const resolved = expectedStackOrder.map((tier) => expectedTierValues[tier]);
    expect(resolved).toEqual([...resolved].sort((a, b) => a - b));
  });

  it('keeps each visible overlay on its expected layer', () => {
    for (const [selector, tier] of overlayExpectations) {
      const block = cssBlockFor(overlayCss, selector);
      expect(block, selector).toContain(
        `z-index: var(--fc-z-${tier}, ${expectedTierValues[tier]});`,
      );
    }
  });

  it('keeps every overlay z-index fallback aligned with the global tier defaults', () => {
    const matches = [...overlayCss.matchAll(zIndexVarPattern)];
    expect(matches.length).toBeGreaterThanOrEqual(overlayExpectations.length);

    for (const [, tier, fallback] of matches) {
      expect(tier, `unknown z-index tier --fc-z-${tier}`).toBeOneOf(expectedStackOrder);
      expect(Number(fallback), `fallback for --fc-z-${tier}`).toBe(
        expectedTierValues[tier as ZTier],
      );
    }
  });
});
