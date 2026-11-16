import { afterEach, describe, expect, it } from 'vitest';
import { mount } from '../../../src/api/spreadsheet/instance.js';
import type { ViewportOptions } from '../../../src/interact/navigation-policy.js';
import { mutators } from '../../../src/store/store.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { pressTab, tabOutOf, tabStops } from '../../test-utils/tab-order.js';

const key = (target: Element, init: KeyboardEventInit): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
};

let sheet: MountedStubSheet;
afterEach(() => {
  sheet.dispose();
  document.body.replaceChildren();
});

const mountOn = (platform: 'default' | 'mac') =>
  mountStubSheet({ locale: 'en', toolbar: true, ui: { platform } });

describe('the sheet bar', () => {
  it.each(['default', 'mac'] as const)(
    'lets Tab and Shift+Tab leave it on the %s profile',
    async (platform) => {
      sheet = await mountOn(platform);
      const bar = sheet.instance.host.querySelector<HTMLElement>('.fc-host__sheetbar');
      if (!bar) throw new Error('No sheet bar.');
      expect(tabOutOf(bar, false, 10)).not.toBeNull();
      expect(tabOutOf(bar, true, 10)).not.toBeNull();
    },
  );
});

const KEYS = [
  { key: 'Delete' },
  { key: 'x' },
  { key: 'ArrowDown' },
  { key: 'Enter' },
  { key: 'PageDown' },
];

// A key pressed on a bar or pane control in the host belongs to that control, not to the grid.
describe('controls outside the grid', () => {
  it.each(['default', 'mac'] as const)(
    'leave the cells and the active cell alone on the %s profile',
    async (platform) => {
      sheet = await mountOn(platform);
      const store = sheet.instance.store;
      const host = sheet.instance.host;
      const controls = tabStops().filter(
        (el) =>
          host.contains(el) &&
          el !== host &&
          !el.closest('.fc-host__grid') &&
          !el.matches('input, textarea, select'),
      );
      expect(controls.some((el) => el.closest('.fc-host__ribbon'))).toBe(true);
      const reached: string[] = [];
      const reset = (): void => {
        sheet.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 7);
        mutators.replaceCells(store, sheet.workbook.cells(0));
        mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
      };
      reset();
      for (const control of controls) {
        for (const init of KEYS) {
          if (!control.isConnected) break;
          control.focus();
          key(control, init);
          const state = store.getState();
          if (
            state.ui.editor.kind !== 'idle' ||
            state.selection.active.row !== 1 ||
            state.selection.active.col !== 1 ||
            sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 }).kind !== 'number'
          ) {
            reached.push(`${control.className}: ${init.key}`);
            key(document.activeElement ?? host, { key: 'Escape' });
            reset();
          }
        }
      }
      expect(reached).toEqual([]);
    },
  );
});

describe('the grid as the sheet tab stop', () => {
  const mountViewport = (viewport: ViewportOptions) =>
    mountStubSheet({ locale: 'en', toolbar: true, viewport });

  const gridOf = (): HTMLElement => {
    const grid = sheet.instance.host.querySelector<HTMLElement>('.fc-host__grid');
    if (!grid) throw new Error('No grid.');
    return grid;
  };

  it('leaves the mount element out of the tab order and makes the grid the one sheet stop', async () => {
    sheet = await mountOn('default');
    const host = sheet.instance.host;
    const grid = gridOf();
    expect(host.hasAttribute('tabindex')).toBe(false);
    expect(tabStops()).not.toContain(host);
    expect(tabStops().filter((el) => el.matches('[role="grid"]') || grid.contains(el))).toEqual([
      grid,
    ]);
  });

  it('focuses the grid from instance.focus()', async () => {
    installDomStubs();
    const { host, cleanup } = createHostElement();
    const instance = await mount(host, { locale: 'en' });
    try {
      instance.focus();
      expect(document.activeElement).toBe(host.querySelector('.fc-host__grid'));
    } finally {
      instance.dispose();
      cleanup();
      uninstallDomStubs();
    }
  });

  it("keeps Tab on the grid and moves the active cell with tabBoundary 'stop'", async () => {
    sheet = await mountViewport({ range: 'A1:B1', tabBoundary: 'stop' });
    const store = sheet.instance.store;
    const grid = gridOf();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    grid.focus();
    expect(pressTab(false)).toBe(false);
    expect(store.getState().selection.active.col).toBe(1);
    expect(document.activeElement).toBe(grid);
    // At the last cell the press is swallowed and focus stays.
    expect(pressTab(false)).toBe(false);
    expect(document.activeElement).toBe(grid);
  });

  it("moves focus out of the grid from the edge cells with tabBoundary 'leave'", async () => {
    sheet = await mountViewport({ range: 'A1:B1', tabBoundary: 'leave' });
    const store = sheet.instance.store;
    const grid = gridOf();
    const stops = tabStops();
    const index = stops.indexOf(grid);
    expect(index).toBeGreaterThan(0);
    expect(index).toBeLessThan(stops.length - 1);

    mutators.setActive(store, { sheet: 0, row: 0, col: 1 });
    grid.focus();
    expect(pressTab(false)).toBe(true);
    expect(document.activeElement).toBe(stops[index + 1]);

    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    grid.focus();
    expect(pressTab(true)).toBe(true);
    expect(document.activeElement).toBe(stops[index - 1]);
  });
});
