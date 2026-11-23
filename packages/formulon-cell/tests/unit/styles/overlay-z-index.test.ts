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
  base: 16_777_000,
  drawing: 16_777_004,
  grid: 16_777_010,
  dialog: 16_777_020,
  tooltip: 16_777_030,
  popover: 16_777_040,
  callout: 16_777_050,
  menu: 16_777_060,
  error: 16_777_070,
} as const;

type ZTier = Exclude<keyof typeof expectedTierValues, 'base'>;

const expectedStackOrder: ZTier[] = [
  'drawing',
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
  ['.fc-objects', 'grid'],
  ['.fc-find', 'grid'],
  ['.fc-slicer', 'grid'],
  ['.fc-fmtdlg', 'dialog'],
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

  it('stacks every drawing panel in one root, over the sheet and under the in-grid panels', () => {
    // One stacking context at the drawing tier holds the chart, picture and shape
    // panels; the inner roots form none, so the panels interleave by z-index.
    const drawings = cssBlockFor(overlayCss, '.fc-drawings');
    expect(drawings).toContain(`z-index: var(--fc-z-drawing, ${expectedTierValues.drawing});`);
    expect(drawings).toContain('pointer-events: none;');
    expect(drawings).not.toMatch(/overflow|clip|contain:/);
    expect(cssBlockFor(overlayCss, '.fc-charts')).not.toMatch(
      /z-index:|transform:|opacity:|filter:|isolation:/,
    );
    const surfaceCss = walkCss(resolve(styleRoot, 'surface')).join('\n');
    expect(cssBlockFor(surfaceCss, '.fc-host__grid')).not.toMatch(/z-index:/);
    expect(cssBlockFor(surfaceCss, '.fc-host__canvas')).not.toMatch(/z-index:/);
    for (const selector of ['.fc-find', '.fc-quick', '.fc-slicer']) {
      expect(cssBlockFor(overlayCss, selector), selector).toContain(
        `z-index: var(--fc-z-grid, ${expectedTierValues.grid});`,
      );
    }
    expect(cssBlockFor(overlayCss, '.fc-ctxmenu')).toContain(
      `z-index: var(--fc-z-menu, ${expectedTierValues.menu});`,
    );
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

  it('keeps every z-index number an exact float32 integer, so CSS minifiers keep it', () => {
    // Minifiers print numbers through float32; above 2^24 an integer can round to a lower value.
    const css = [baseCss, overlayCss].join('\n');
    const numbers = [
      ...[...css.matchAll(/z-index:[^;]*/g)].flatMap((m) => m[0].match(/\d+/g) ?? []),
      ...[...css.matchAll(/--fc-z-base:\s*(\d+)/g)].map((m) => m[1] ?? ''),
    ].map(Number);
    expect(numbers.length).toBeGreaterThan(0);
    for (const value of numbers) {
      expect(Math.fround(value), String(value)).toBe(value);
      expect(value, String(value)).toBeLessThanOrEqual(2 ** 24);
    }
  });

  it('layers ribbon dropdowns over the in-grid surfaces that share the host stacking context', () => {
    // The ribbon mounts inside the isolated .fc-host, so its menus compete with the
    // Quick Analysis launcher, in-grid panels and drawings by tier, not by a local number.
    expect(cssBlockFor(overlayCss, '.fc-quick__button')).toContain(
      `z-index: var(--fc-z-grid, ${expectedTierValues.grid});`,
    );
    const toolbarCss = walkCss(resolve(styleRoot, '../toolbar')).join('\n');
    const macCss = readFileSync(resolve(styleRoot, '../platform-mac.css'), 'utf8');
    const popover = `var(--fc-z-popover, ${expectedTierValues.popover})`;
    for (const selector of [
      '.fc-tb__submenu',
      '.fc-tb__rb-dd__list',
      '.fc-tb__merge-menu',
      '.fc-tb__command-menu',
      '.fc-tb__ribbon-shell--peek > .fc-tb__ribbon:not([hidden])',
    ]) {
      expect(cssBlockFor(toolbarCss, selector), selector).toContain(`z-index: ${popover};`);
    }
    for (const selector of [
      '.fc-tb__menu.fc-tb__tablestyle-ctx',
      '.fc-tb__ribbon-shell--peek .fc-tb__ribbon-display',
    ]) {
      expect(cssBlockFor(toolbarCss, selector), selector).toContain(
        `z-index: calc(${popover} + 1);`,
      );
    }
    expect(cssBlockFor(macCss, '> .fc-tb__ribbon-tools')).toContain(`z-index: ${popover};`);
    expect([toolbarCss, macCss].join('\n')).not.toMatch(/z-index:\s*(?:40|80|9\d);/);
  });
});
