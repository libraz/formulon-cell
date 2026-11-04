import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { attachBorderDraw } from '../../../src/interact/border-draw.js';
import { attachFormatPainter } from '../../../src/interact/format-painter.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const fire = (host: HTMLElement, type: string, x: number, y: number): PointerEvent => {
  const e = new PointerEvent(type, {
    clientX: x,
    clientY: y,
    button: 0,
    bubbles: true,
    cancelable: true,
    pointerId: 1,
  });
  host.dispatchEvent(e);
  return e;
};

describe('pointer gestures that are cancelled', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let captureThrows = false;

  beforeEach(() => {
    captureThrows = false;
    host = document.createElement('div');
    const captured = new Set<number>();
    Object.assign(host, {
      setPointerCapture(id: number) {
        if (captureThrows) throw new Error('NotFoundError');
        captured.add(id);
      },
      releasePointerCapture(id: number) {
        captured.delete(id);
      },
      hasPointerCapture(id: number) {
        return captured.has(id);
      },
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
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('format painter does not apply formats on pointercancel and stays armed', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    const handle = attachFormatPainter({ host, store, wb: () => null });
    handle.activate();
    fire(host, 'pointerdown', 300, 100);
    fire(host, 'pointercancel', 300, 100);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }))).toBe(
      undefined,
    );
    expect(handle.isActive()).toBe(true);
    // A later stray move must not extend the selection any more.
    const before = store.getState().selection.range;
    fire(host, 'pointermove', 500, 200);
    expect(store.getState().selection.range).toBe(before);
    handle.detach();
  });

  it('format painter survives a throwing setPointerCapture', () => {
    captureThrows = true;
    const handle = attachFormatPainter({ host, store, wb: () => null });
    handle.activate();
    expect(() => fire(host, 'pointerdown', 300, 100)).not.toThrow();
    handle.detach();
  });

  it('format painter skips selection writes while the hit cell is unchanged', () => {
    const handle = attachFormatPainter({ host, store, wb: () => null });
    handle.activate();
    fire(host, 'pointerdown', 300, 100);
    fire(host, 'pointermove', 310, 105);
    const spy = vi.fn();
    const off = store.subscribe(spy);
    fire(host, 'pointermove', 312, 106);
    fire(host, 'pointermove', 314, 107);
    expect(spy).not.toHaveBeenCalled();
    off();
    handle.detach();
  });

  it('border draw does not write borders on pointercancel', () => {
    const handle = attachBorderDraw({ host, store });
    handle.activate('grid', 'thin');
    fire(host, 'pointerdown', 300, 100);
    fire(host, 'pointercancel', 400, 150);
    expect(store.getState().format.formats.size).toBe(0);
    handle.detach();
  });

  it('border draw survives a throwing setPointerCapture', () => {
    captureThrows = true;
    const handle = attachBorderDraw({ host, store });
    handle.activate('grid', 'thin');
    expect(() => fire(host, 'pointerdown', 300, 100)).not.toThrow();
    handle.detach();
  });
});
