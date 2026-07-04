import { afterEach, describe, expect, it } from 'vitest';

import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: the overlay portal is the single themed container that all
 * body-attached floating UI mounts into. It must be created on mount, track
 * the host theme, and be torn down on dispose so no orphan container leaks
 * into `<body>`.
 */
describe('integration: overlay portal lifecycle', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => sheet?.dispose());

  const portals = () => document.querySelectorAll<HTMLElement>('.fc-overlay-portal');

  it('creates one themed portal in <body> on mount', async () => {
    sheet = await mountStubSheet({ theme: 'ink' });
    const list = Array.from(portals());
    expect(list).toHaveLength(1);
    const portal = list[0];
    expect(portal?.parentElement).toBe(document.body);
    expect(portal?.dataset.fcTheme).toBe('ink');
  });

  it('keeps the portal theme in sync with setTheme()', async () => {
    sheet = await mountStubSheet({ theme: 'paper' });
    expect(portals()[0]?.dataset.fcTheme).toBe('paper');

    sheet.instance.setTheme('ink');
    expect(portals()[0]?.dataset.fcTheme).toBe('ink');
  });

  it('removes the portal from <body> on dispose()', async () => {
    sheet = await mountStubSheet();
    expect(portals()).toHaveLength(1);

    sheet.instance.dispose();
    expect(portals()).toHaveLength(0);
  });
});
