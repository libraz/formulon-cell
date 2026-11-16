import { afterEach, describe, expect, it } from 'vitest';
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
  if (active.closest('.fc-host__ribbon')) return 'ribbon';
  if (active.matches('.fc-host__formulabar-tag')) return 'nameBox';
  if (active === sheetFocusTarget(host)) return 'grid';
  if (active.closest('.fc-host__sheetbar-tabs')) return 'sheetTab';
  if (active.closest('.fc-host__statusbar')) return 'statusBar';
  if (active.closest('[data-fc-surface="pane"]')) return 'pane';
  return 'other';
};

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
});
