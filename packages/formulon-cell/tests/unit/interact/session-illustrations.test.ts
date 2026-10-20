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
      labels: { picture: 'Picture overlay', shape: 'Shape overlay' },
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
    expect(overlays[0]?.querySelectorAll('.fc-drawing-handle')).toHaveLength(8);

    handle.detach();
  });

  it('names pictures and shapes the way Excel does, numbered among their kind', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertIllustration(store, {
      id: 'image-without-alt',
      kind: 'image',
      sheet: 0,
      src: 'data:image/png;base64,image',
    });
    mutators.upsertIllustration(store, { id: 'oval-1', kind: 'shape', shape: 'oval', sheet: 0 });
    mutators.upsertIllustration(store, { id: 'oval-2', kind: 'shape', shape: 'oval', sheet: 0 });
    mutators.upsertIllustration(store, {
      id: 'shape-without-kind',
      kind: 'shape',
      sheet: 0,
    });

    const handle = attachSessionIllustrations({ host, store });
    const overlays = Array.from(host.querySelectorAll<HTMLElement>('.fc-illustration'));

    expect(overlays.map((overlay) => overlay.getAttribute('aria-label'))).toEqual([
      'Picture 1',
      'Oval 1',
      'Oval 2',
      'Rectangle 1',
    ]);
    handle.setLabels({ ovalName: '楕円 {n}' });
    expect(
      host.querySelectorAll<HTMLElement>('.fc-illustration')[2]?.getAttribute('aria-label'),
    ).toBe('楕円 2');

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
    expect(panel.querySelector('.fc-drawing-handle')).toBeNull();

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
    expect(overlay?.getAttribute('aria-roledescription')).toBe('Shape');
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

  const freeShape = (
    store: ReturnType<typeof createSpreadsheetStore>,
    patch: Partial<Parameters<typeof mutators.upsertIllustration>[1]> = {},
  ): void => {
    mutators.upsertIllustration(store, {
      id: 'shape-1',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
      x: 20,
      y: 30,
      w: 160,
      h: 96,
      ...patch,
    });
  };

  it('selects like Excel: names itself, and leaves on an outside press or Escape', () => {
    const host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    freeShape(store, { shape: 'oval' });
    const handle = attachSessionIllustrations({ host, store });
    const overlay = (): HTMLElement => {
      const found = host.querySelector<HTMLElement>('.fc-illustration');
      if (!found) throw new Error('missing session illustration overlay');
      return found;
    };

    firePointer(overlay(), 'pointerdown');
    expect(overlay().classList.contains('fc-illustration--selected')).toBe(true);
    expect(store.getState().ui.selectedObject).toEqual({ id: 'shape-1', name: 'Oval 1' });
    firePointer(window, 'pointerup');

    firePointer(host, 'pointerdown');
    expect(overlay().classList.contains('fc-illustration--selected')).toBe(false);
    expect(store.getState().ui.selectedObject).toBeNull();

    firePointer(overlay(), 'pointerdown');
    firePointer(window, 'pointerup');
    const escaped = fireKey(overlay(), 'Escape');
    expect(escaped.defaultPrevented).toBe(true);
    expect(store.getState().ui.selectedObject).toBeNull();
    expect(document.activeElement).toBe(host);

    handle.detach();
    host.remove();
  });

  it('moves on a drag past the threshold and records the drag as one change', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    freeShape(store);
    const handle = attachSessionIllustrations({ host, store, history });
    const overlay = host.querySelector<HTMLElement>('.fc-illustration');
    if (!overlay) throw new Error('missing session illustration overlay');

    firePointer(overlay, 'pointerdown', { clientX: 100, clientY: 100 });
    firePointer(window, 'pointermove', { clientX: 101, clientY: 101 });
    firePointer(window, 'pointerup', { clientX: 101, clientY: 101 });
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ x: 20, y: 30 });
    expect(history.canUndo()).toBe(false);

    const panel = host.querySelector<HTMLElement>('.fc-illustration') ?? overlay;
    firePointer(panel, 'pointerdown', { clientX: 100, clientY: 100 });
    firePointer(window, 'pointermove', { clientX: 110, clientY: 105 });
    firePointer(window, 'pointermove', { clientX: 140, clientY: 120 });
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ x: 20, y: 30 });
    expect(panel.style.left).toBe('60px');
    firePointer(window, 'pointerup', { clientX: 140, clientY: 120 });
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ x: 60, y: 50 });
    expect(history.undo()).toBe(true);
    expect(store.getState().illustrations.illustrations[0]).toMatchObject({ x: 20, y: 30 });
    expect(history.canUndo()).toBe(false);

    handle.detach();
    host.remove();
  });

  it('resizes from any handle; a picture keeps its shape, a shape only with Shift', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    freeShape(store);
    mutators.upsertIllustration(store, {
      id: 'image-1',
      kind: 'image',
      sheet: 0,
      src: 'data:image/png;base64,image',
      x: 300,
      y: 30,
      w: 200,
      h: 100,
    });
    const handle = attachSessionIllustrations({ host, store });
    const grip = (id: string, name: string): HTMLElement => {
      const found = host.querySelector<HTMLElement>(
        `[data-illustration-id="${id}"] .fc-drawing-handle[data-handle="${name}"]`,
      );
      if (!found) throw new Error(`missing ${name} handle`);
      return found;
    };
    const drag = (target: HTMLElement, dx: number, dy: number, shiftKey = false): void => {
      firePointer(target, 'pointerdown', { clientX: 0, clientY: 0 });
      firePointer(window, 'pointermove', { clientX: dx, clientY: dy, shiftKey });
      firePointer(window, 'pointerup', { clientX: dx, clientY: dy, shiftKey });
    };
    const item = (id: string) =>
      store.getState().illustrations.illustrations.find((entry) => entry.id === id);

    drag(grip('shape-1', 'nw'), -20, -10);
    expect(item('shape-1')).toMatchObject({ x: 0, y: 20, w: 180, h: 106 });
    drag(grip('shape-1', 'se'), 180, 0);
    expect(item('shape-1')).toMatchObject({ w: 360, h: 106 });
    drag(grip('shape-1', 'se'), 360, 0, true);
    expect(item('shape-1')).toMatchObject({ w: 720, h: 212 });

    drag(grip('image-1', 'se'), 200, 0);
    expect(item('image-1')).toMatchObject({ x: 300, y: 30, w: 400, h: 200 });
    drag(grip('image-1', 'e'), -100, 0);
    expect(item('image-1')).toMatchObject({ w: 300, h: 200 });

    handle.detach();
    host.remove();
  });

  it('leaves the name box to a chart selected by another overlay', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    freeShape(store);
    const handle = attachSessionIllustrations({ host, store });
    mutators.setSelectedObject(store, { id: 'chart-1', name: 'Chart 1' });
    firePointer(host, 'pointerdown');
    handle.detach();
    expect(store.getState().ui.selectedObject).toEqual({ id: 'chart-1', name: 'Chart 1' });
    host.remove();
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
