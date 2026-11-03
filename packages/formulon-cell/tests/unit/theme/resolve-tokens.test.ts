import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

import { type ResolvedTheme, resolveTheme } from '../../../src/theme/resolve.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const TOKEN_KEYS: (keyof ResolvedTheme)[] = [
  'bg',
  'bgRail',
  'bgElev',
  'bgHeader',
  'fg',
  'fgMute',
  'fgFaint',
  'fgStrong',
  'rule',
  'ruleStrong',
  'accent',
  'accentFg',
  'accentSoft',
  'cellErrorFg',
  'cellFormulaFg',
  'cellBoolFg',
  'cellNumFg',
  'hoverStripe',
  'headerFg',
  'headerFgActive',
  'fontUi',
  'fontMono',
  'textCell',
  'textHeader',
];

describe('theme/resolve — token completeness', () => {
  let host: HTMLElement | undefined;

  afterEach(() => {
    host?.remove();
    host = undefined;
  });

  it('produces every required key when nothing is set (fallbacks fire)', () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    const theme = resolveTheme(host);
    for (const key of TOKEN_KEYS) {
      expect(theme[key], `missing key ${key}`).not.toBeUndefined();
    }
  });

  it('returns string color tokens and numeric text-size tokens', () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    const theme = resolveTheme(host);
    expect(typeof theme.bg).toBe('string');
    expect(typeof theme.accent).toBe('string');
    expect(typeof theme.textCell).toBe('number');
    expect(typeof theme.textHeader).toBe('number');
  });

  it('parses numeric text tokens (default 13 / 11.5 when no CSS)', () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    const theme = resolveTheme(host);
    expect(theme.textCell).toBe(13);
    expect(theme.textHeader).toBe(11.5);
  });

  it('honors inline custom-property overrides on the host', () => {
    host = document.createElement('div');
    host.style.setProperty('--fc-bg', '#101010');
    host.style.setProperty('--fc-accent', '#abcdef');
    document.body.appendChild(host);
    const theme = resolveTheme(host);
    expect(theme.bg).toBe('#101010');
    expect(theme.accent).toBe('#abcdef');
  });

  it('survives empty-string custom property values by falling back', () => {
    host = document.createElement('div');
    host.style.setProperty('--fc-bg', '');
    document.body.appendChild(host);
    const theme = resolveTheme(host);
    expect(theme.bg).toBe('#faf7f1');
  });

  it('keeps the paper grid chrome close to Excel 365 light headers', () => {
    const css = readFileSync(join(root, 'src/styles/theme-paper.css'), 'utf8');
    expect(css).toContain('--fc-bg-header: #e6e6e6;');
    expect(css).toContain('--fc-rule: #d9d9d9;');
    expect(css).toContain('--fc-rule-strong: #bfbfbf;');
    expect(css).toContain('--fc-accent: #107c41;');
    expect(css).toContain('--fc-accent-strong: #0b5f31;');
    expect(css).toContain('--fc-cell-auto-fg: #000000;');
    expect(css).toContain('--fc-cell-bool-fg: var(--fc-cell-auto-fg);');
    expect(css).toContain('--fc-header-fg-active: #107c41;');
  });

  it('keeps the paper bottom chrome on a neutral desktop Excel surface', () => {
    const css = readFileSync(join(root, 'src/styles/theme-paper.css'), 'utf8');
    expect(css).toContain('--fc-sheetbar-tab-hover: #e9e9e9;');
    expect(css).toContain('--fc-sheetbar-tab-rule: #d9d9d9;');
    expect(css).toContain('--fc-statusbar-bg: #f3f2f1;');
    expect(css).toContain('--fc-statusbar-border: #d9d9d9;');
    expect(css).toContain('--fc-statusbar-fg: #605e5c;');
    expect(css).toContain('--fc-statusbar-fg-strong: #201f1e;');
    expect(css).toContain('--fc-statusbar-control: #c8c6c4;');
  });
});

describe('theme/resolve — header strip shading', () => {
  it('stays flat in the header fill unless both ends are set', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    host.style.setProperty('--fc-header-bg', '#abcdef');
    expect(resolveTheme(host)).toMatchObject({
      headerStripOuter: '#abcdef',
      headerStripInner: '#abcdef',
    });
    host.style.setProperty('--fc-header-strip-outer', '#fdfdfd');
    host.style.setProperty('--fc-header-strip-inner', '#f7f7f7');
    expect(resolveTheme(host)).toMatchObject({
      headerStripOuter: '#fdfdfd',
      headerStripInner: '#f7f7f7',
    });
    host.remove();
  });

  it('keeps header separators flat in the header rule unless the outer end is set', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    host.style.setProperty('--fc-header-rule', '#dddcda');
    expect(resolveTheme(host)).toMatchObject({ headerRule: '#dddcda', headerRuleOuter: '#dddcda' });
    host.style.setProperty('--fc-header-rule-outer', '#fbfbfb');
    expect(resolveTheme(host)).toMatchObject({ headerRuleOuter: '#fbfbfb' });
    host.remove();
  });
});

describe('theme/resolve — selected header tokens', () => {
  it('defaults to an accent bar and a semibold label', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    host.style.setProperty('--fc-accent', '#123456');
    const theme = resolveTheme(host);
    expect(theme.headerActiveBar).toBe('#123456');
    expect(theme.headerFgActiveWeight).toBe(600);
    host.remove();
  });

  it('reads the bar colour and the label weight from their tokens', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    host.style.setProperty('--fc-header-active-bar', 'transparent');
    host.style.setProperty('--fc-header-fg-active-weight', '400');
    const theme = resolveTheme(host);
    expect(theme.headerActiveBar).toBe('transparent');
    expect(theme.headerFgActiveWeight).toBe(400);
    host.remove();
  });
});
