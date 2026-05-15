import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { inheritHostTokens } from '../../../src/interact/inherit-host-tokens.js';

describe('inheritHostTokens', () => {
  let host: HTMLElement;
  let target: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    target = document.createElement('div');
    document.body.appendChild(host);
    document.body.appendChild(target);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('copies the well-known custom properties from host computed style onto target inline style', () => {
    host.style.setProperty('--fc-bg', '#101010');
    host.style.setProperty('--fc-fg', '#f0f0f0');
    host.style.setProperty('--fc-rule', '#404040');
    host.style.setProperty('--fc-accent', '#1e90ff');

    inheritHostTokens(host, target);

    expect(target.style.getPropertyValue('--fc-bg')).toBe('#101010');
    expect(target.style.getPropertyValue('--fc-fg')).toBe('#f0f0f0');
    expect(target.style.getPropertyValue('--fc-rule')).toBe('#404040');
    expect(target.style.getPropertyValue('--fc-accent')).toBe('#1e90ff');
  });

  it('skips tokens absent on the host (does not stamp empty strings)', () => {
    // Only one token set; everything else should remain unset on target.
    host.style.setProperty('--fc-bg', '#222');
    inheritHostTokens(host, target);
    expect(target.style.getPropertyValue('--fc-bg')).toBe('#222');
    // No spurious empty-string stamps for tokens the host doesn't have.
    expect(target.style.getPropertyValue('--fc-accent')).toBe('');
    expect(target.style.getPropertyValue('--fc-fg')).toBe('');
    expect(target.style.getPropertyValue('--fc-rule-strong')).toBe('');
  });

  it('forwards the color-scheme property when the host sets it', () => {
    host.style.setProperty('color-scheme', 'dark');
    inheritHostTokens(host, target);
    expect(target.style.getPropertyValue('color-scheme')).toBe('dark');
  });

  it('re-reading after the host changes a token reflects the new value (idempotent re-call)', () => {
    host.style.setProperty('--fc-accent', '#aaa');
    inheritHostTokens(host, target);
    expect(target.style.getPropertyValue('--fc-accent')).toBe('#aaa');

    host.style.setProperty('--fc-accent', '#bbb');
    inheritHostTokens(host, target);
    expect(target.style.getPropertyValue('--fc-accent')).toBe('#bbb');
  });

  it('does not touch unrelated CSS properties on target', () => {
    target.style.color = 'rgb(255, 0, 0)';
    target.style.padding = '4px';
    host.style.setProperty('--fc-bg', '#000');
    inheritHostTokens(host, target);
    expect(target.style.color).toBe('rgb(255, 0, 0)');
    expect(target.style.padding).toBe('4px');
    expect(target.style.getPropertyValue('--fc-bg')).toBe('#000');
  });

  it('forwards radius and font tokens too', () => {
    host.style.setProperty('--fc-radius-sm', '4px');
    host.style.setProperty('--fc-radius-md', '8px');
    host.style.setProperty('--fc-font-ui', 'Inter');
    host.style.setProperty('--fc-font-mono', 'JetBrains Mono');
    inheritHostTokens(host, target);
    expect(target.style.getPropertyValue('--fc-radius-sm')).toBe('4px');
    expect(target.style.getPropertyValue('--fc-radius-md')).toBe('8px');
    expect(target.style.getPropertyValue('--fc-font-ui')).toBe('Inter');
    expect(target.style.getPropertyValue('--fc-font-mono')).toBe('JetBrains Mono');
  });
});
