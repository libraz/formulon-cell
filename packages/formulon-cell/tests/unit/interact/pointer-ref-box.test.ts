import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
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

// Default test layout: header 52 x 30, cells 104 x 28. A1:A3 spans x 52..156, y 30..114.
describe('formula-edit reference box drag', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let detach: () => void;
  const replaceReferenceAt = vi.fn();

  beforeEach(async () => {
    replaceReferenceAt.mockReset();
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
      ui: {
        ...s.ui,
        editor: { kind: 'edit', raw: '=SUM(A1:A3)', caret: 11 },
        editorRefs: [{ r0: 0, c0: 0, r1: 2, c1: 0, colorIndex: 0 }],
      },
    }));
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    detach = attachPointer(host, store, wb, undefined, null, () => ({
      isFormulaEdit: () => true,
      editedCell: () => ({ sheet: 0, row: 0, col: 1 }),
      insertRefAtCaret: vi.fn(),
      replaceReferenceAt,
    }));
  });

  afterEach(() => {
    detach();
    document.body.innerHTML = '';
  });

  it('moves the box by the cell delta between press and pointer', () => {
    const down = fire(host, 'pointerdown', 155, 70);
    expect(down.defaultPrevented).toBe(true);
    fire(host, 'pointermove', 363, 70);
    expect(replaceReferenceAt).toHaveBeenNthCalledWith(1, 0, 'C1:C3', { phase: 'start' });
    fire(host, 'pointermove', 470, 70);
    expect(replaceReferenceAt).toHaveBeenNthCalledWith(2, 0, 'E1:E3', { phase: 'update' });
    fire(host, 'pointerup', 470, 70);
  });

  it('resizes from the corner handle keeping the opposite corner', () => {
    fire(host, 'pointerdown', 154, 112);
    fire(host, 'pointermove', 100, 170);
    expect(replaceReferenceAt).toHaveBeenCalledWith(0, 'A1:A6', { phase: 'start' });
    fire(host, 'pointerup', 100, 170);
  });

  it('uses a resize cursor over a corner and the arrow over an edge', () => {
    fire(host, 'pointermove', 154, 112);
    expect(host.style.cursor).toBe('nwse-resize');
    fire(host, 'pointermove', 155, 70);
    expect(host.style.cursor).toBe('default');
  });

  it('inserts nothing when a cell inside the box is pressed', () => {
    fire(host, 'pointerdown', 100, 70);
    fire(host, 'pointermove', 100, 100);
    expect(replaceReferenceAt).not.toHaveBeenCalled();
  });
});

describe('right-button drag of the selection frame', () => {
  it('raises the drop menu request for the dropped range', async () => {
    const host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
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
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.extendRangeTo(store, { sheet: 0, row: 2, col: 2 });
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const detach = attachPointer(host, store, wb);
    const seen = vi.fn();
    host.addEventListener('fc:movemenu', seen);
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
    press('pointerdown', 200, 58);
    press('pointermove', 408, 144);
    expect(store.getState().ui.movePreview).toEqual({ sheet: 0, r0: 4, c0: 3, r1: 5, c1: 4 });
    press('pointerup', 408, 144);
    expect(seen).toHaveBeenCalledTimes(1);
    const [event] = seen.mock.calls[0] as [CustomEvent];
    const detail = event.detail;
    expect(detail.src).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    expect(detail.dest).toEqual({ sheet: 0, r0: 4, c0: 3, r1: 5, c1: 4 });
    expect(store.getState().selection.range).toEqual(detail.src);
    detach();
    host.remove();
  });
});
