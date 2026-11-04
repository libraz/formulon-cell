import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ensureMacInk, type MacInkController } from '../../../src/interact/mac-ink.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const firePointer = (target: EventTarget, type: string, x: number, y: number): void => {
  target.dispatchEvent(
    new PointerEvent(type, {
      bubbles: true,
      cancelable: true,
      pointerId: 1,
      button: 0,
      clientX: x,
      clientY: y,
    }),
  );
};

describe('Mac ink rendering cost', () => {
  let host: HTMLElement;
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let controller: MacInkController | undefined;

  beforeEach(() => {
    host = document.createElement('div');
    host.dataset.fcPlatform = 'mac';
    grid = document.createElement('div');
    grid.className = 'fc-host__grid';
    Object.defineProperty(grid, 'getBoundingClientRect', {
      configurable: true,
      value: () => ({
        left: 0,
        top: 0,
        right: 500,
        bottom: 300,
        width: 500,
        height: 300,
        x: 0,
        y: 0,
        toJSON: () => ({}),
      }),
    });
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, widthPx: 500 } }));
  });

  afterEach(() => {
    controller?.dispose();
    vi.restoreAllMocks();
    document.body.innerHTML = '';
  });

  it('does not rebuild the display SVG when only the selection changes', () => {
    controller = ensureMacInk({ host, store });
    const display = host.querySelector('.fc-mac-ink__display') as SVGElement;
    const defs = display.querySelector('defs');
    mutators.setActive(store, { sheet: 0, row: 4, col: 3 });
    expect(display.querySelector('defs')).toBe(defs);
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart: 5 } }));
    expect(display.querySelector('defs')).not.toBe(defs);
  });

  it('redraws the live stroke once per frame', () => {
    const frames: FrameRequestCallback[] = [];
    vi.spyOn(window, 'requestAnimationFrame').mockImplementation((cb) => frames.push(cb));
    controller = ensureMacInk({ host, store });
    controller?.activate('pen-red');
    const input = host.querySelector('.fc-mac-ink__input') as SVGElement;
    firePointer(input, 'pointerdown', 48, 30);
    frames.length = 0;
    const before = input.querySelector('defs');
    firePointer(input, 'pointermove', 60, 36);
    firePointer(input, 'pointermove', 72, 40);
    firePointer(input, 'pointermove', 84, 42);
    expect(frames).toHaveLength(1);
    expect(input.querySelector('defs')).toBe(before);
    frames[0]?.(0);
    expect(input.querySelector('defs')).not.toBe(before);
    expect(input.querySelector('path')).toBeTruthy();
  });
});
