import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const move = (host: HTMLElement, x: number, y: number): void => {
  host.dispatchEvent(
    new PointerEvent('pointermove', { clientX: x, clientY: y, bubbles: true, pointerId: 1 }),
  );
};

// Default test layout: header 52 x 30, cells 104 x 28.
describe('Mac grid cursors', () => {
  let outer: HTMLElement;
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let detach: () => void;

  beforeEach(async () => {
    outer = document.createElement('div');
    outer.className = 'fc-host';
    outer.dataset.fcPlatform = 'mac';
    host = document.createElement('div');
    outer.appendChild(host);
    document.body.appendChild(outer);
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
    detach = attachPointer(host, store, await WorkbookHandle.createDefault({ preferStub: true }));
  });

  afterEach(() => {
    detach();
    outer.remove();
  });

  it('uses the cell cross over a cell body', () => {
    move(host, 200, 70);
    expect(host.style.cursor).toBe('cell');
  });

  it('uses black arrows over the headers', () => {
    move(host, 100, 10);
    expect(decodeURIComponent(host.style.cursor)).toContain('M7 3h6v7h3.5L10 17');
    move(host, 10, 40);
    expect(decodeURIComponent(host.style.cursor)).toContain('M3 7h7V3.5');
  });

  it('uses split triangles on the resize boundaries', () => {
    move(host, 154, 10);
    expect(decodeURIComponent(host.style.cursor)).toContain('M1.5 10 7 5v10');
    move(host, 10, 56);
    expect(decodeURIComponent(host.style.cursor)).toContain('M10 1.5 15 7H5z');
  });
});
