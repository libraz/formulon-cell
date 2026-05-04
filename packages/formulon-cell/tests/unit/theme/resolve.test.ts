import { afterEach, describe, expect, it } from 'vitest';
import { resolveTheme } from '../../../src/theme/resolve.js';

const mountHost = (cssText = ''): HTMLElement => {
  const host = document.createElement('div');
  host.style.cssText = cssText;
  document.body.appendChild(host);
  return host;
};

describe('resolveTheme', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('returns built-in fallbacks when no CSS variables are set', () => {
    const host = mountHost();
    const theme = resolveTheme(host);
    expect(theme.bg).toBe('#faf7f1');
    expect(theme.fg).toBe('#15171c');
    expect(theme.accent).toBe('#d83a14');
    expect(theme.fontUi).toBe('system-ui, sans-serif');
    expect(theme.fontMono).toBe('ui-monospace, monospace');
    // num() falls back to its default when the CSS value parses as NaN.
    expect(theme.textCell).toBe(13);
    expect(theme.textHeader).toBe(11.5);
  });

  it('reads custom properties from the host computed style', () => {
    const host = mountHost(
      '--fc-bg: #000000; --fc-accent: #ff00ff; --fc-text-cell: 18px; --fc-text-header: 14px;',
    );
    const theme = resolveTheme(host);
    expect(theme.bg).toBe('#000000');
    expect(theme.accent).toBe('#ff00ff');
    expect(theme.textCell).toBe(18);
    expect(theme.textHeader).toBe(14);
  });

  it('falls back to defaults when a numeric custom property is non-numeric', () => {
    const host = mountHost('--fc-text-cell: not-a-number;');
    const theme = resolveTheme(host);
    expect(theme.textCell).toBe(13);
  });

  it('exposes every documented field on the resolved theme', () => {
    const host = mountHost();
    const theme = resolveTheme(host);
    const expected: Array<keyof typeof theme> = [
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
    for (const k of expected) expect(theme[k]).toBeDefined();
  });
});
