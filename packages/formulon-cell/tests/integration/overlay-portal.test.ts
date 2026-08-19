import { afterEach, describe, expect, it } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../src/mount.js';
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

  it('keeps the portal alive when a superseded instance disposes late', async () => {
    // Remounting the same host — an async mount() racing React StrictMode's
    // double-effect — leaves the first instance disposing after the second has
    // taken over. Its teardown must not take the live instance's overlays with
    // it.
    sheet = await mountStubSheet();
    const first = sheet.instance;
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const second = await Spreadsheet.mount(sheet.host, { workbook });

    first.dispose();
    const portal = portals()[0];
    expect(portal).toBeDefined();
    expect(portal?.querySelector('.fc-ctxmenu')).not.toBeNull();

    second.dispose();
    workbook.dispose();
    expect(portals()).toHaveLength(0);
  });
});
