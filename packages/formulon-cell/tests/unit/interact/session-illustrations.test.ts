import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  arrangeSessionIllustration,
  clearSessionIllustration,
  duplicateSessionIllustration,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { cellRectUnclamped, layoutForView } from '../../../src/render/geometry.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const firePointer = (
  target: EventTarget,
  type: string,
  init: PointerEventInit = {},
): PointerEvent => {
  const e = new PointerEvent(type, {
    bubbles: true,
    pointerId: 1,
    button: 0,
    clientX: 0,
    clientY: 0,
    ...init,
  });
  target.dispatchEvent(e);
  return e;
};

const fireKey = (target: EventTarget, key: string, init: KeyboardEventInit = {}): KeyboardEvent => {
  const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key, ...init });
  target.dispatchEvent(e);
  return e;
};

describe('attachSessionIllustrations', () => {
  it('renders session image and shape overlays for the active sheet only', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertIllustration(store, {
      id: 'image-1',
      kind: 'image',
      sheet: 0,
      src: 'data:image/png;base64,image',
      alt: 'Image label',
    });
    mutators.upsertIllustration(store, {
      id: 'shape-1',
      kind: 'shape',
      shape: 'oval',
      sheet: 0,
      color: '#107c10',
    });
    mutators.upsertIllustration(store, {
      id: 'shape-other-sheet',
      kind: 'shape',
      shape: 'arrow',
      sheet: 1,
    });

    const handle = attachSessionIllustrations({
      host,
      store,
      pictureLabel: 'Picture overlay',
      shapeLabel: 'Shape overlay',
      resizeLabel: 'Resize illustration',
    });

    const overlays = Array.from(host.querySelectorAll<HTMLElement>('.fc-illustration'));
    expect(overlays.map((overlay) => overlay.dataset.illustrationId)).toEqual([
      'image-1',
      'shape-1',
    ]);
    expect(host.querySelector<HTMLImageElement>('img')?.alt).toBe('Image label');
    expect(host.querySelector('svg ellipse[stroke="#107c10"]')).toBeTruthy();
    expect(overlays[0]?.getAttribute('aria-roledescription')).toBe('Picture overlay');
    expect(overlays[1]?.getAttribute('aria-roledescription')).toBe('Shape overlay');
    expect(host.querySelector('[aria-label="Resize illustration"]')).toBeTruthy();

    handle.detach();
  });

  it('does not expose English fallback labels when no overlay labels are supplied', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertIllustration(store, {
      id: 'image-without-alt',
      kind: 'image',
      sheet: 0,
      src: 'data:image/png;base64,image',
    });
    mutators.upsertIllustration(store, {
      id: 'shape-without-kind',
      kind: 'shape',
      sheet: 0,
    });

    const handle = attachSessionIllustrations({ host, store });
    const overlays = Array.from(host.querySelectorAll<HTMLElement>('.fc-illustration'));

    expect(overlays.map((overlay) => overlay.getAttribute('aria-label'))).toEqual([
      'image-without-alt',
      'shape-without-kind',
    ]);
    expect(host.querySelector('[aria-label="Resize shape"]')).toBeNull();
    expect(host.textContent).not.toContain('Picture');
    expect(host.textContent).not.toContain('Shape');

    handle.detach();
  });

  it('tracks native anchor offsets and dimensions through zoom and scroll', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, zoom: 2, rowStart: 3, colStart: 2 },
    }));
    mutators.upsertIllustration(store, {
      id: 'imported-native',
      kind: 'image',
      sheet: 0,
      src: 'data:image/png;base64,image',
      nativeObjectId: 7,
      nativeGeometryEditable: false,
      nativeAnchor: {
        anchorKind: 0,
        editAs: 1,
        fromRow: 5,
        fromCol: 6,
        fromRowOff: 10 * 9525,
        fromColOff: 12 * 9525,
        toRow: 0,
        toCol: 0,
        toRowOff: 0,
        toColOff: 0,
        cx: 100 * 9525,
        cy: 50 * 9525,
      },
    });

    const handle = attachSessionIllustrations({ host, store });
    const panel = host.querySelector<HTMLElement>('.fc-illustration');
    if (!panel) throw new Error('native overlay missing');
    const first = cellRectUnclamped(
      layoutForView(store.getState()),
      store.getState().viewport,
      5,
      6,
    );
    expect(Number.parseFloat(panel.style.left)).toBeCloseTo(first.x + 12 * 2);
    expect(Number.parseFloat(panel.style.top)).toBeCloseTo(first.y + 10 * 2);
    expect(Number.parseFloat(panel.style.width)).toBeCloseTo(100 * 2);
    expect(Number.parseFloat(panel.style.height)).toBeCloseTo(50 * 2);
    expect(panel.querySelector('[role="separator"]')).toBeNull();

    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, rowStart: 4, colStart: 3 },
    }));
    const repositioned = host.querySelector<HTMLElement>('.fc-illustration');
    if (!repositioned) throw new Error('native overlay missing after scroll');
    const second = cellRectUnclamped(
      layoutForView(store.getState()),
      store.getState().viewport,
      5,
      6,
    );
    expect(Number.parseFloat(repositioned.style.left)).toBeCloseTo(second.x + 12 * 2);
    expect(Number.parseFloat(repositioned.style.top)).toBeCloseTo(second.y + 10 * 2);
    handle.detach();
  });

  it('moves, resizes, deletes, and records keyboard changes through shared commands', () => {
    const host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertIllustration(store, {
      id: 'shape-1',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
      x: 20,
      y: 30,
      w: 160,
      h: 96,
    });

    const handle = attachSessionIllustrations({ host, store, history });
    const overlay = host.querySelector<HTMLElement>('.fc-illustration');
    expect(overlay?.tabIndex).toBe(0);
    expect(overlay?.getAttribute('aria-roledescription')).toBe('shape');
    if (!overlay) throw new Error('missing session illustration overlay');

    const move = fireKey(overlay, 'ArrowRight');
    expect(move.defaultPrevented).toBe(true);
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ x: 28, y: 30 });
    expect(history.undo()).toBe(true);
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ x: 20, y: 30 });
    expect(history.redo()).toBe(true);
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ x: 28, y: 30 });

    const resize = fireKey(
      host.querySelector<HTMLElement>('.fc-illustration') ?? overlay,
      'ArrowDown',
      {
        shiftKey: true,
      },
    );
    expect(resize.defaultPrevented).toBe(true);
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ w: 160, h: 104 });

    fireKey(host.querySelector<HTMLElement>('.fc-illustration') ?? overlay, 'Delete');
    expect(store.getState().illustrations.illustrations).toHaveLength(0);
    expect(host.querySelector('.fc-illustration')).toBeNull();
    expect(document.activeElement).toBe(host);
    expect(history.undo()).toBe(true);
    expect(store.getState().illustrations.illustrations.map((item) => item.id)).toEqual([
      'shape-1',
    ]);

    handle.detach();
  });

  it('uses shared protection gates while moving or deleting overlays', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertIllustration(store, {
      id: 'shape-1',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
      x: 20,
      y: 30,
      w: 160,
      h: 96,
    });
    setProtectedSheet(store, 0, true);

    const handle = attachSessionIllustrations({ host, store });
    const overlay = host.querySelector<HTMLElement>('.fc-illustration');
    if (!overlay) throw new Error('missing session illustration overlay');

    firePointer(overlay, 'pointerdown', { clientX: 10, clientY: 10 });
    firePointer(window, 'pointermove', { clientX: 50, clientY: 50 });
    firePointer(window, 'pointerup', { clientX: 50, clientY: 50 });
    fireKey(host.querySelector<HTMLElement>('.fc-illustration') ?? overlay, 'Delete');

    expect(store.getState().illustrations.illustrations).toMatchObject([
      {
        id: 'shape-1',
        x: 20,
        y: 30,
      },
    ]);

    handle.detach();
  });

  it('reauthorizes illustration mutations while a viewer policy is active', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history,
    });
    const unregister = registerInteractionController(store, controller);
    controller.setPolicy(viewerPolicy());
    mutators.upsertIllustration(store, {
      id: 'viewer-image',
      kind: 'image',
      sheet: 0,
      src: 'data:image/png;base64,image',
      x: 20,
      y: 30,
      w: 160,
      h: 96,
    });
    mutators.upsertIllustration(store, {
      id: 'viewer-shape',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
      x: 40,
      y: 50,
      w: 160,
      h: 96,
    });

    expect(updateSessionIllustration(store, 'viewer-image', { x: 80 }, history, wb)).toBeNull();
    expect(clearSessionIllustration(store, 'viewer-image', history, wb)).toBe(false);
    expect(duplicateSessionIllustration(store, 'viewer-image', history, wb)).toBeNull();
    expect(arrangeSessionIllustration(store, 'viewer-image', 'bring-front', history)).toBeNull();
    expect(store.getState().illustrations.illustrations.map((item) => item.id)).toEqual([
      'viewer-image',
      'viewer-shape',
    ]);
    expect(history.canUndo()).toBe(false);

    unregister();
    controller.dispose();
    wb.dispose();
    host.remove();
  });
});
