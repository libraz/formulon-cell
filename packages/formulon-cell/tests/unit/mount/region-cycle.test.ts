import { afterEach, describe, expect, it } from 'vitest';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { sheetFocusTarget } from '../../test-utils/sheet-focus.js';

let sheet: MountedStubSheet;
afterEach(() => {
  sheet.dispose();
  document.body.replaceChildren();
});

const f6 = (target: Element, shiftKey = false): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', {
    key: 'F6',
    shiftKey,
    bubbles: true,
    cancelable: true,
  });
  target.dispatchEvent(event);
  return event;
};

const stopOf = (host: HTMLElement): string => {
  const active = document.activeElement;
  if (!active) return 'none';
  if (active.closest('.fc-host__ribbon, [data-external-ribbon]')) return 'ribbon';
  if (active.matches('.fc-host__formulabar-tag')) return 'nameBox';
  if (active === sheetFocusTarget(host)) return 'grid';
  if (active.closest('.fc-host__sheetbar-tabs')) return 'sheetTab';
  if (active.closest('.fc-host__statusbar')) return 'statusBar';
  if (active.closest('[data-fc-surface="pane"]')) return 'pane';
  return 'other';
};

const nextFrame = (): Promise<void> =>
  new Promise((resolve) => requestAnimationFrame(() => resolve()));

const ring = (host: HTMLElement, shiftKey: boolean, presses: number): string[] => {
  const seen: string[] = [];
  for (let i = 0; i < presses; i += 1) {
    f6(document.activeElement ?? host, shiftKey);
    seen.push(stopOf(host));
  }
  return seen;
};

describe.each(['default', 'mac'] as const)('F6 region cycling on the %s profile', (platform) => {
  const mountFull = async (toolbar: boolean) => {
    sheet = await mountStubSheet({ locale: 'en', toolbar, ui: { platform } });
    sheet.instance.openWatchWindow();
    sheetFocusTarget(sheet.host).focus();
  };

  it('walks ribbon, name box, grid, sheet tab, status bar and pane, then wraps', async () => {
    await mountFull(true);
    expect(ring(sheet.host, false, 7)).toEqual([
      'sheetTab',
      'statusBar',
      'pane',
      'ribbon',
      'nameBox',
      'grid',
      'sheetTab',
    ]);
  });

  it('moves exactly one stop on F6 with the Watch Window open, and nothing pulls focus after', async () => {
    await mountFull(true);
    f6(document.activeElement as Element);
    expect(stopOf(sheet.host)).toBe('sheetTab');
    await nextFrame();
    await nextFrame();
    expect(stopOf(sheet.host)).toBe('sheetTab');
  });

  it('does not move focus into the Watch Window when the store changes', async () => {
    await mountFull(true);
    const store = sheet.instance.store;
    const grid = sheetFocusTarget(sheet.host);
    mutators.addWatch(store, { sheet: 0, row: 0, col: 0 });
    await nextFrame();
    expect(document.activeElement).toBe(grid);
    // Closed and reopened through the store, not through the panel's own open().
    mutators.setWatchPanelOpen(store, false);
    mutators.setWatchPanelOpen(store, true);
    await nextFrame();
    await nextFrame();
    expect(document.activeElement).toBe(grid);
  });

  it('runs the ring in exact reverse on Shift+F6', async () => {
    await mountFull(true);
    expect(ring(sheet.host, true, 7)).toEqual([
      'nameBox',
      'ribbon',
      'pane',
      'statusBar',
      'sheetTab',
      'grid',
      'nameBox',
    ]);
  });

  it('skips the ribbon when none is mounted', async () => {
    await mountFull(false);
    expect(ring(sheet.host, false, 6)).toEqual([
      'sheetTab',
      'statusBar',
      'pane',
      'nameBox',
      'grid',
      'sheetTab',
    ]);
  });

  it('leaves F6 to the page when the grid is the only region', async () => {
    sheet = await mountStubSheet({
      locale: 'en',
      ui: {
        platform,
        features: { formulaBar: false, sheetTabs: false, statusBar: false, watchWindow: false },
      },
    });
    const grid = sheetFocusTarget(sheet.host);
    grid.focus();
    for (const shiftKey of [false, true]) {
      const event = f6(grid, shiftKey);
      expect(event.defaultPrevented).toBe(false);
      expect(document.activeElement).toBe(grid);
    }
  });

  it('does nothing inside an open modal dialog', async () => {
    await mountFull(true);
    sheet.instance.openGoTo();
    const active = document.activeElement as HTMLElement;
    expect(active.closest('[data-fc-surface="modal"]')).not.toBeNull();
    const event = f6(active);
    expect(event.defaultPrevented).toBe(false);
    expect(document.activeElement).toBe(active);
  });

  describe('with the ribbon mounted outside the sheet', () => {
    let ribbonHost: HTMLElement;
    let toolbar: ReturnType<typeof Spreadsheet.mountToolbar>;
    const mountExternal = async () => {
      await mountFull(false);
      ribbonHost = document.createElement('div');
      ribbonHost.dataset.externalRibbon = '';
      sheet.host.before(ribbonHost);
      toolbar = Spreadsheet.mountToolbar(ribbonHost, sheet.instance, {});
      sheetFocusTarget(sheet.host).focus();
    };

    it('walks the same ring as an in-mount ribbon, forward', async () => {
      await mountExternal();
      expect(ring(sheet.host, false, 7)).toEqual([
        'sheetTab',
        'statusBar',
        'pane',
        'ribbon',
        'nameBox',
        'grid',
        'sheetTab',
      ]);
    });

    it('walks the same ring in reverse', async () => {
      await mountExternal();
      expect(ring(sheet.host, true, 7)).toEqual([
        'nameBox',
        'ribbon',
        'pane',
        'statusBar',
        'sheetTab',
        'grid',
        'nameBox',
      ]);
    });

    it('moves on from inside the external ribbon in both directions', async () => {
      await mountExternal();
      expect(toolbar.focusActiveTab()).toBe(true);
      f6(document.activeElement as Element);
      expect(stopOf(sheet.host)).toBe('nameBox');
      expect(toolbar.focusActiveTab()).toBe(true);
      f6(document.activeElement as Element, true);
      expect(stopOf(sheet.host)).toBe('pane');
    });

    it('enters at the active ribbon tab', async () => {
      await mountExternal();
      toolbar.setActiveTab('view');
      f6(sheetFocusTarget(sheet.host), true);
      f6(document.activeElement as Element, true);
      expect(stopOf(sheet.host)).toBe('ribbon');
      expect((document.activeElement as HTMLElement).dataset.ribbonTab).toBe('view');
    });

    it('skips a hidden external ribbon', async () => {
      await mountExternal();
      ribbonHost.hidden = true;
      expect(ring(sheet.host, true, 2)).toEqual(['nameBox', 'pane']);
    });

    it('drops the stop once the toolbar is disposed', async () => {
      await mountExternal();
      toolbar.dispose();
      expect(ring(sheet.host, false, 4)).toEqual(['sheetTab', 'statusBar', 'pane', 'nameBox']);
      expect(ring(sheet.host, true, 1)).toEqual(['pane']);
    });
  });
});
