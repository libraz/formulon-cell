import { afterEach, describe, expect, it } from 'vitest';
import { resolveFlags } from '../../../src/extensions/features.js';
import { ja } from '../../../src/i18n/strings.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createMountChrome } from '../../../src/mount/chrome.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

describe('locale switcher chrome slot', () => {
  let host: HTMLElement | undefined;
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    if (host) disposeOverlayPortal(host);
    host = undefined;
    sheet?.dispose();
    sheet = undefined;
    document.body.replaceChildren();
  });

  it('places the compact row after the optional viewbar and before the workspace', () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    const chrome = createMountChrome({
      host,
      getStrings: () => ja,
      flags: resolveFlags({
        formulaBar: false,
        viewToolbar: true,
        localeSwitcher: true,
        sheetTabs: false,
        statusBar: true,
        watchWindow: false,
      }),
      onSheetTabContextMenu: () => {},
    });

    expect(chrome.localeSwitcher.parentElement).toBe(host);
    expect(chrome.localeSwitcher.previousElementSibling).toBe(chrome.viewbar);
    expect(chrome.localeSwitcher.nextElementSibling).toBe(chrome.sheetWorkspace);
    expect(chrome.statusbar.parentElement).toBe(host);
    expect(chrome.statusbar.previousElementSibling).toBe(chrome.a11yLive);
  });

  it('keeps ordering stable when viewbar and locale switcher are toggled live', () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    const chrome = createMountChrome({
      host,
      getStrings: () => ja,
      flags: resolveFlags({
        formulaBar: false,
        viewToolbar: false,
        localeSwitcher: true,
        sheetTabs: false,
        statusBar: false,
        watchWindow: false,
      }),
      onSheetTabContextMenu: () => {},
    });

    chrome.setChromeAttached('viewbar', true);
    expect(chrome.viewbar.nextElementSibling).toBe(chrome.localeSwitcher);
    expect(chrome.localeSwitcher.nextElementSibling).toBe(chrome.sheetWorkspace);

    chrome.setChromeAttached('localeSwitcher', false);
    expect(chrome.viewbar.nextElementSibling).toBe(chrome.sheetWorkspace);
    chrome.setChromeAttached('localeSwitcher', true);
    expect(chrome.viewbar.nextElementSibling).toBe(chrome.localeSwitcher);
    expect(chrome.localeSwitcher.nextElementSibling).toBe(chrome.sheetWorkspace);
  });

  it('adds and removes the feature node through the live feature API', async () => {
    sheet = await mountStubSheet({
      features: { localeSwitcher: true, viewToolbar: false },
    });
    const selector = (): HTMLSelectElement | null =>
      sheet?.host.querySelector<HTMLSelectElement>('[data-fc-locale-select]') ?? null;

    expect(selector()).not.toBeNull();
    sheet.instance.setFeatures({ localeSwitcher: false, viewToolbar: false });
    expect(selector()).toBeNull();
    sheet.instance.setFeatures({ localeSwitcher: true, viewToolbar: false });
    expect(selector()).not.toBeNull();
  });
});
