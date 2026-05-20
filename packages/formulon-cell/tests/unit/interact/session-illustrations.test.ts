import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
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
});
