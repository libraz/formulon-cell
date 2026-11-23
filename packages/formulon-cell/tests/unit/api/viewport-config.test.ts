import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { mount, type SpreadsheetInstance } from '../../../src/api/spreadsheet/instance.js';
import type { ViewportConfig } from '../../../src/api/spreadsheet/types.js';
import { isNavigationAddrAllowed } from '../../../src/interact/navigation-policy.js';
import { spreadsheetInternals } from '../../../src/mount/internals.js';
import { mutators } from '../../../src/store/store.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';
import { sheetFocusTarget } from '../../test-utils/sheet-focus.js';

const at = (row: number, col: number) => ({ sheet: 0, row, col });

describe('viewport option', () => {
  let instance: SpreadsheetInstance | undefined;
  let host: HTMLElement;
  let cleanup: () => void;

  beforeEach(() => {
    installDomStubs();
    ({ host, cleanup } = createHostElement());
  });
  afterEach(() => {
    instance?.dispose();
    instance = undefined;
    cleanup();
    uninstallDomStubs();
  });

  const mountWith = async (viewport: ViewportConfig) => {
    instance = await mount(host, { viewport });
    return spreadsheetInternals(instance).store;
  };

  const tab = (): KeyboardEvent => {
    const e = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    sheetFocusTarget(host).dispatchEvent(e);
    return e;
  };

  it('limits a fixed range to its selectable list', async () => {
    const store = await mountWith({ mode: 'fixed', range: 'A1:H7', selectable: ['B2:C3'] });
    expect(isNavigationAddrAllowed(store, at(1, 1))).toBe(true);
    expect(isNavigationAddrAllowed(store, at(0, 0))).toBe(false);
    expect(isNavigationAddrAllowed(store, at(4, 4))).toBe(false);
  });

  it('passes a selectable predicate the cell as a context object', async () => {
    const store = await mountWith({
      mode: 'fixed',
      range: 'A1:H7',
      selectable: ({ addr }) => addr.row !== 1,
    });
    expect(isNavigationAddrAllowed(store, at(0, 3))).toBe(true);
    expect(isNavigationAddrAllowed(store, at(1, 3))).toBe(false);
  });

  it('limits the whole sheet to a selectable list in auto mode', async () => {
    const store = await mountWith({ mode: 'auto', selectable: ['B2:C3'] });
    expect(isNavigationAddrAllowed(store, at(2, 2))).toBe(true);
    expect(isNavigationAddrAllowed(store, at(100, 100))).toBe(false);
  });

  it('keeps Tab inside the range at its last stop by default', async () => {
    const store = await mountWith({ mode: 'fixed', range: 'A1:B1' });
    mutators.setActive(store, at(0, 1));
    expect(tab().defaultPrevented).toBe(true);
  });

  it("lets Tab leave the grid at the last stop with tabBoundary 'leave'", async () => {
    const store = await mountWith({ mode: 'fixed', range: 'A1:B1', tabBoundary: 'leave' });
    mutators.setActive(store, at(0, 1));
    expect(tab().defaultPrevented).toBe(false);
  });

  it('takes the same object at mount and through update', async () => {
    const viewport: ViewportConfig = {
      mode: 'fixed',
      range: 'A1:H7',
      selectable: ['B2:C3'],
      tabBoundary: 'leave',
    };
    const atMount = await mountWith(viewport);
    const mounted = [at(1, 1), at(0, 0)].map((addr) => isNavigationAddrAllowed(atMount, addr));
    expect(instance?.viewport.options).toEqual(viewport);

    instance?.update({ viewport: null });
    expect(isNavigationAddrAllowed(atMount, at(0, 0))).toBe(true);
    instance?.update({ viewport });
    const updated = [at(1, 1), at(0, 0)].map((addr) => isNavigationAddrAllowed(atMount, addr));
    expect(updated).toEqual(mounted);
    expect(updated).toEqual([true, false]);
    expect(instance?.viewport.options).toEqual(viewport);
  });
});

describe('viewport.scrollTo', () => {
  let instance: SpreadsheetInstance;
  let cleanup: () => void;

  beforeEach(async () => {
    installDomStubs();
    let host: HTMLElement;
    ({ host, cleanup } = createHostElement());
    instance = await mount(host);
  });
  afterEach(() => {
    instance.dispose();
    cleanup();
    uninstallDomStubs();
  });

  it('brings an off-screen target into view, not only the active cell', () => {
    const store = spreadsheetInternals(instance).store;
    mutators.setViewportSize(store, 10, 5);
    instance.viewport.scrollTo('A26');
    const s = store.getState();
    expect(s.selection.active).toMatchObject({ row: 25, col: 0 });
    expect(s.viewport.rowStart).toBeLessThanOrEqual(25);
    expect(s.viewport.rowStart + s.viewport.rowCount - 1).toBeGreaterThanOrEqual(25);
  });
});
