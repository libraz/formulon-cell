import { afterEach, describe, expect, it, vi } from 'vitest';

import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: `mount({ toolbar: true })` builds the ribbon toolbar inside the
 * host in a single call, exposes it as `instance.toolbar`, and tears it down on
 * dispose. The shell lives under the themed `.fc-host`, so grid and toolbar
 * inherit the same `data-fc-theme` by cascade.
 */
describe('integration: single-call ribbon toolbar', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  const shellOf = (host: HTMLElement) => host.querySelector<HTMLElement>('.fc-tb__ribbon-shell');

  it('mounts the ribbon inside the host and exposes instance.toolbar', async () => {
    sheet = await mountStubSheet({ toolbar: true, theme: 'ink' });
    const { host, instance } = sheet;

    expect(instance.toolbar).not.toBeNull();
    expect(typeof instance.toolbar?.getActiveTab).toBe('function');

    const shell = shellOf(host);
    expect(shell).toBeTruthy();
    // The shell descends from the themed host, so `[data-fc-theme]` tokens
    // reach the toolbar through the cascade — no separate theme attribute.
    expect(shell?.closest('.fc-host')).toBe(host);
    expect(host.dataset.fcTheme).toBe('ink');

    // The ribbon host is a `display: contents` first child so the shell is the
    // first flex item in the host column, above the formula bar.
    const ribbonHost = host.querySelector<HTMLElement>('.fc-host__ribbon');
    expect(ribbonHost).toBe(host.firstElementChild);
    expect(ribbonHost?.style.display).toBe('contents');
  });

  it('leaves instance.toolbar null when toolbar is not requested', async () => {
    sheet = await mountStubSheet();
    expect(sheet.instance.toolbar).toBeNull();
    expect(shellOf(sheet.host)).toBeNull();
  });

  it('forwards a toolbar options object while keeping the single-call defaults', async () => {
    const onTabChange = vi.fn();
    sheet = await mountStubSheet({ toolbar: { activeTab: 'insert', onTabChange } });
    const tb = sheet.instance.toolbar;
    expect(tb).not.toBeNull();

    // Host-supplied options reach mountToolbar…
    expect(tb?.getActiveTab()).toBe('insert');
    // …and the single-call defaults survive the merge (dynamicDropdowns: true
    // is applied under the spread, so the dropdown api is wired).
    expect(tb?.dropdownsApi).not.toBeNull();

    // Lifecycle callbacks passed through fire on the matching state change.
    tb?.setActiveTab('home');
    expect(onTabChange).toHaveBeenCalledWith('home');
  });

  it('disposes the toolbar and removes the ribbon host on dispose()', async () => {
    sheet = await mountStubSheet({ toolbar: true });
    const { host, instance } = sheet;
    expect(shellOf(host)).toBeTruthy();

    instance.dispose();
    expect(host.querySelector('.fc-host__ribbon')).toBeNull();
    expect(shellOf(host)).toBeNull();
  });
});
