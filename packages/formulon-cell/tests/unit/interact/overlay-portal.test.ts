import { afterEach, describe, expect, it, vi } from 'vitest';

import {
  disposeOverlayPortal,
  ensureOverlayPortal,
  hasOpenLocaleDraft,
  overlayPortalFor,
  refreshOverlayPortal,
  setOverlayOptions,
} from '../../../src/interact/overlay-portal.js';

const hosts: HTMLElement[] = [];

function makeHost(parent: HTMLElement = document.body): HTMLElement {
  const host = document.createElement('div');
  host.className = 'fc-host';
  parent.appendChild(host);
  hosts.push(host);
  return host;
}

afterEach(() => {
  for (const host of hosts.splice(0)) disposeOverlayPortal(host);
  document.body.replaceChildren();
  Reflect.deleteProperty(document, 'fullscreenElement');
  vi.restoreAllMocks();
});

describe('overlay portal', () => {
  it.each(['dialog', 'alertdialog'])(
    'blocks visible %s drafts only in their owning sheet',
    (role) => {
      const host = makeHost();
      const other = makeHost();
      const portal = ensureOverlayPortal(host);
      const modal = document.createElement('div');
      modal.setAttribute('role', role);
      modal.setAttribute('aria-modal', 'true');
      portal.append(modal);
      expect(hasOpenLocaleDraft(host)).toBe(true);
      expect(hasOpenLocaleDraft(other)).toBe(false);
      modal.hidden = true;
      expect(hasOpenLocaleDraft(host)).toBe(false);
      modal.hidden = false;
      portal.style.display = 'none';
      expect(hasOpenLocaleDraft(host)).toBe(false);
      portal.style.display = '';
      modal.remove();
      expect(hasOpenLocaleDraft(host)).toBe(false);
    },
  );

  it.each(['fc-find', 'fc-cmtnote', 'fc-filter-dropdown'])(
    'preserves editable nonmodal %s drafts while allowing other panels',
    (className) => {
      const host = makeHost();
      const panel = document.createElement('div');
      panel.setAttribute('role', 'dialog');
      panel.setAttribute('aria-modal', 'false');
      const portal = ensureOverlayPortal(host);
      portal.append(panel);
      expect(hasOpenLocaleDraft(host)).toBe(false);
      panel.className = className;
      expect(hasOpenLocaleDraft(host)).toBe(true);
      panel.hidden = true;
      expect(hasOpenLocaleDraft(host)).toBe(false);
    },
  );

  it('uses the host owner document body and preserves portal identity and theme', () => {
    const host = makeHost();
    host.dataset.fcTheme = 'ink';
    host.lang = 'fr';

    const first = ensureOverlayPortal(host);
    const second = ensureOverlayPortal(host);

    expect(second).toBe(first);
    expect(first.parentElement).toBe(document.body);
    expect(first.ownerDocument).toBe(document);
    expect(first.dataset.fcTheme).toBe('ink');
    expect(first.lang).toBe('fr');
    host.lang = 'de';
    refreshOverlayPortal(host);
    expect(first.lang).toBe('de');
  });

  it('moves the owned portal between explicit roots without removing host roots', () => {
    const firstRoot = document.createElement('section');
    const secondRoot = document.createElement('section');
    document.body.append(firstRoot, secondRoot);
    const host = makeHost();
    const portal = ensureOverlayPortal(host, { root: firstRoot });
    const child = document.createElement('button');
    portal.appendChild(child);

    setOverlayOptions(host, { root: secondRoot });

    expect(portal.parentElement).toBe(secondRoot);
    expect(portal.querySelector('button')).toBe(child);
    expect(firstRoot.isConnected).toBe(true);
    expect(secondRoot.isConnected).toBe(true);

    disposeOverlayPortal(host);
    expect(portal.isConnected).toBe(false);
    expect(secondRoot.isConnected).toBe(true);
  });

  it('refreshes resolver roots for overlayPortalFor and keeps cached children', () => {
    const firstRoot = document.createElement('section');
    const secondRoot = document.createElement('section');
    document.body.append(firstRoot, secondRoot);
    const host = makeHost();
    let root = firstRoot;
    const portal = ensureOverlayPortal(host, { root: () => root });
    const trigger = document.createElement('button');
    portal.appendChild(trigger);

    root = secondRoot;
    expect(overlayPortalFor(trigger)).toBe(portal);
    expect(portal.parentElement).toBe(secondRoot);
    expect(portal.contains(trigger)).toBe(true);
  });

  it('keeps an explicit root ahead of dialog and fullscreen defaults', () => {
    const dialog = document.createElement('dialog');
    const explicit = document.createElement('section');
    document.body.append(dialog, explicit);
    dialog.setAttribute('open', '');
    const host = makeHost(dialog);
    let fullscreen: Element | null = null;
    Object.defineProperty(document, 'fullscreenElement', {
      configurable: true,
      get: () => fullscreen,
    });
    const fullscreenRoot = document.createElement('section');
    dialog.appendChild(fullscreenRoot);
    fullscreen = fullscreenRoot;

    const portal = ensureOverlayPortal(host, { root: explicit });
    expect(portal.parentElement).toBe(explicit);
  });

  it('defaults to an open native dialog and returns to body when it closes', () => {
    const dialog = document.createElement('dialog');
    dialog.setAttribute('open', '');
    document.body.appendChild(dialog);
    const host = makeHost(dialog);
    const portal = ensureOverlayPortal(host);

    expect(portal.parentElement).toBe(dialog);

    dialog.removeAttribute('open');
    refreshOverlayPortal(host);
    expect(portal.parentElement).toBe(document.body);
  });

  it('follows fullscreenchange for the containing fullscreen element', () => {
    const fullscreenRoot = document.createElement('section');
    document.body.appendChild(fullscreenRoot);
    const host = makeHost();
    let fullscreen: Element | null = null;
    Object.defineProperty(document, 'fullscreenElement', {
      configurable: true,
      get: () => fullscreen,
    });
    const portal = ensureOverlayPortal(host);
    expect(portal.parentElement).toBe(document.body);

    fullscreen = fullscreenRoot;
    fullscreenRoot.appendChild(host);
    document.dispatchEvent(new Event('fullscreenchange'));
    expect(portal.parentElement).toBe(fullscreenRoot);

    fullscreen = null;
    fullscreenRoot.removeChild(host);
    document.dispatchEvent(new Event('fullscreenchange'));
    expect(portal.parentElement).toBe(document.body);
  });

  it('uses a host owner document for portals and rejects a foreign root', () => {
    const otherDocument = document.implementation.createHTMLDocument('other');
    const otherHost = otherDocument.createElement('div');
    otherHost.className = 'fc-host';
    otherDocument.body.appendChild(otherHost);
    hosts.push(otherHost);
    const portal = ensureOverlayPortal(otherHost);

    expect(portal.ownerDocument).toBe(otherDocument);
    expect(portal.parentElement).toBe(otherDocument.body);
    expect(() => ensureOverlayPortal(otherHost, { root: document.createElement('div') })).toThrow(
      /ownerDocument/,
    );
    expect(portal.parentElement).toBe(otherDocument.body);
  });

  it('rejects invalid resolver roots without leaving a portal or losing a valid root', () => {
    const host = makeHost();
    expect(() =>
      ensureOverlayPortal(host, {
        root: () => null as unknown as HTMLElement,
      }),
    ).toThrow(/HTMLElement/);
    expect(document.querySelector('.fc-overlay-portal')).toBeNull();

    const validRoot = document.createElement('section');
    document.body.appendChild(validRoot);
    const portal = ensureOverlayPortal(host, { root: validRoot });
    expect(() =>
      setOverlayOptions(host, {
        root: () => null as unknown as HTMLElement,
      }),
    ).toThrow(/HTMLElement/);
    expect(portal.parentElement).toBe(validRoot);
  });

  it('removes its fullscreen listener and honors the superseded-instance guard', () => {
    const removeListener = vi.spyOn(document, 'removeEventListener');
    const host = makeHost();
    host.dataset.fcInstId = 'new';
    const portal = ensureOverlayPortal(host);

    disposeOverlayPortal(host, 'old');
    expect(portal.isConnected).toBe(true);

    disposeOverlayPortal(host, 'new');
    expect(portal.isConnected).toBe(false);
    expect(removeListener).toHaveBeenCalledWith('fullscreenchange', expect.any(Function));
  });
});
