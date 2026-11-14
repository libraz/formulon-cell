import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ja } from '../../../src/i18n/strings.js';
import { attachStatusBar } from '../../../src/interact/status-bar.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const seed = (platform: 'mac' | null, stale: { value: boolean }) => {
  const host = document.createElement('div');
  host.className = 'fc-host';
  if (platform) host.dataset.fcPlatform = platform;
  const bar = document.createElement('div');
  bar.className = 'fc-host__statusbar';
  host.appendChild(bar);
  document.body.appendChild(host);
  const onRecalc = vi.fn(() => {
    stale.value = false;
  });
  const handle = attachStatusBar({
    statusbar: bar,
    store: createSpreadsheetStore(),
    getEngineLabel: () => 'engine',
    strings: ja,
    hasStaleValues: () => stale.value,
    onRecalc,
  });
  return { bar, handle, onRecalc };
};

const button = (bar: HTMLElement): HTMLButtonElement | null =>
  bar.querySelector<HTMLButtonElement>('.fc-host__statusbar-recalc');

describe('status bar Calculate button', () => {
  beforeEach(() => document.body.replaceChildren());
  afterEach(() => document.body.replaceChildren());

  it('shows the bundle label and help on the Mac host while values are stale', () => {
    const { bar } = seed('mac', { value: true });
    const b = button(bar);
    expect(b).not.toBeNull();
    expect(b?.hidden).toBe(false);
    expect(b?.textContent).toBe('再計算');
    expect(b?.title).toBe(ja.mac.chrome.calculateHint);
    expect(b?.title).toContain('F9');
  });

  it('recalculates on press and hides once nothing is stale', () => {
    const { bar, handle, onRecalc } = seed('mac', { value: true });
    button(bar)?.click();
    expect(onRecalc).toHaveBeenCalledTimes(1);
    handle.refresh();
    expect(button(bar)?.hidden).toBe(true);
  });

  it('stays hidden when nothing is stale', () => {
    const { bar } = seed('mac', { value: false });
    expect(button(bar)?.hidden).toBe(true);
  });

  it('is not built on the default profile', () => {
    const { bar } = seed(null, { value: true });
    expect(button(bar)?.hidden ?? true).toBe(true);
  });
});
