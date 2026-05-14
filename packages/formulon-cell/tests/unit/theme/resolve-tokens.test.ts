import { afterEach, describe, expect, it } from 'vitest';

import { type ResolvedTheme, resolveTheme } from '../../../src/theme/resolve.js';

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
});
