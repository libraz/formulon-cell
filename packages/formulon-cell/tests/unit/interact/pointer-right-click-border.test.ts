import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachContextMenu } from '../../../src/interact/context-menu.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

// Header 52 x 30, cells 104 x 28: A1 spans x 52..156, y 30..58.
describe('right-click on the selection border', () => {
  let host: HTMLElement;
  let store: ReturnType<typeof createSpreadsheetStore>;
  let detachPointer: () => void;
  let detachMenu: () => void;

  const press = (type: string, x: number, y: number): void => {
    host.dispatchEvent(
      new PointerEvent(type, {
        clientX: x,
        clientY: y,
        button: 2,
        buttons: type === 'pointerup' ? 0 : 2,
        bubbles: true,
        cancelable: true,
        pointerId: 1,
      }),
    );
  };
  const menuOpen = (): boolean => {
    const root = document.querySelector<HTMLElement>('.fc-ctxmenu');
    return !!root && root.style.display !== 'none';
  };

  beforeEach(async () => {
    host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        defaultColWidth: 104,
        defaultRowHeight: 28,
        headerColWidth: 52,
        headerRowHeight: 30,
      },
    }));
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    detachPointer = attachPointer(host, store, wb);
    detachMenu = attachContextMenu({ getLocale: () => 'en-US', host, store, wb });
  });

  afterEach(() => {
    detachMenu();
    detachPointer();
    host.remove();
  });

  it('opens the cell menu and keeps the selection when released without moving', () => {
    press('pointerdown', 100, 59);
    press('pointerup', 100, 59);
    expect(menuOpen()).toBe(true);
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
  });

  it('still raises the drop menu request after a drag past the threshold', () => {
    const seen = vi.fn();
    host.addEventListener('fc:movemenu', seen);
    press('pointerdown', 100, 59);
    press('pointermove', 100, 120);
    press('pointerup', 100, 120);
    expect(seen).toHaveBeenCalledTimes(1);
    expect(menuOpen()).toBe(false);
  });
});
