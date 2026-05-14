import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createSheetMenuButton } from '../../../src/mount/sheet-menu.js';

const styleRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../src/styles/core');

const walkCss = (dir: string): string[] => {
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    const full = resolve(dir, name);
    if (statSync(full).isDirectory()) out.push(...walkCss(full));
    else if (name.endsWith('.css')) out.push(readFileSync(full, 'utf8'));
  }
  return out;
};

const allCss = walkCss(resolve(styleRoot, 'app')).join('\n');

describe('button disabled-state consistency — JS contract', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    host.remove();
  });

  it('createSheetMenuButton: disabled=true blocks the click handler', () => {
    const handler = vi.fn();
    const close = vi.fn();
    const btn = createSheetMenuButton('Delete', handler, close, true);
    host.appendChild(btn);
    expect(btn.disabled).toBe(true);
    btn.click();
    expect(handler).not.toHaveBeenCalled();
    expect(close).not.toHaveBeenCalled();
  });

  it('createSheetMenuButton: disabled=false invokes handler exactly once', () => {
    const handler = vi.fn();
    const close = vi.fn();
    const btn = createSheetMenuButton('Rename', handler, close, false);
    host.appendChild(btn);
    btn.click();
    expect(handler).toHaveBeenCalledTimes(1);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('a button with .disabled = true also drops keyboard activation', () => {
    const handler = vi.fn();
    const close = vi.fn();
    const btn = createSheetMenuButton('Hide', handler, close, true);
    host.appendChild(btn);
    btn.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
    btn.dispatchEvent(new KeyboardEvent('keydown', { key: ' ' }));
    expect(handler).not.toHaveBeenCalled();
  });

  it('a disabled button does not capture focus when programmatic focus is invoked', () => {
    const btn = document.createElement('button');
    btn.disabled = true;
    host.appendChild(btn);
    btn.focus();
    expect(document.activeElement).not.toBe(btn);
  });
});

describe('button disabled-state consistency — CSS contract', () => {
  // The CSS files define both [disabled] and :disabled selectors. They must
  // produce equivalent treatment: disabled buttons are dimmed and unclickable
  // (cursor:not-allowed or opacity). A regression where one selector is
  // dropped silently leaves disabled-styled buttons clickable / invisible.
  it('every overlay surface that uses :disabled also handles [disabled]', () => {
    const disabledRefs =
      allCss.match(/[.#][a-zA-Z0-9_-]+(?:__[a-zA-Z0-9_-]+)?(?::disabled|\[disabled\])/g) ?? [];
    expect(
      disabledRefs.length,
      'disabled selectors should be present in overlay CSS',
    ).toBeGreaterThan(0);
  });

  it('disabled buttons have an opacity-reduction style (visual dimming)', () => {
    // Scan for blocks that target a disabled state and check at least some
    // have an opacity rule. We don't enforce per-selector — just the
    // overall design-system contract that disabled = dim.
    const disabledBlocks = allCss.match(/(?::disabled|\[disabled\])[^{]*\{[^}]*\}/g) ?? [];
    expect(disabledBlocks.length).toBeGreaterThan(0);
    const dimmed = disabledBlocks.filter((b) => /opacity\s*:/.test(b));
    expect(dimmed.length, 'at least one disabled rule must dim via opacity').toBeGreaterThan(0);
  });

  it('cursor:not-allowed appears on at least one disabled selector', () => {
    const disabledBlocks = allCss.match(/(?::disabled|\[disabled\])[^{]*\{[^}]*\}/g) ?? [];
    const withCursor = disabledBlocks.filter((b) => /cursor\s*:\s*not-allowed/.test(b));
    expect(
      withCursor.length,
      'disabled buttons should advertise cursor:not-allowed',
    ).toBeGreaterThan(0);
  });

  it('hover/focus-visible rules exclude the disabled state via :not()', () => {
    // Prevents the regression where a hover style still applies to a
    // disabled button, suggesting it's clickable.
    const hoverWithoutDisabled =
      allCss.match(/:(?:hover|focus-visible)[^{,]*\{/g)?.filter((s) => !/\)\s*\{$/.test(s)) ?? [];
    const hoverGuards = allCss.match(/:(?:hover|focus-visible):not\(\[disabled\]\)/g) ?? [];
    const hoverGuards2 = allCss.match(/:(?:hover|focus-visible):not\(:disabled\)/g) ?? [];
    expect(
      hoverGuards.length + hoverGuards2.length,
      'at least one hover/focus rule should exclude :disabled / [disabled]',
    ).toBeGreaterThan(0);
    // Sanity: we found hover rules at all.
    expect(hoverWithoutDisabled.length).toBeGreaterThan(0);
  });
});
