import { afterEach, describe, expect, it, vi } from 'vitest';
import { createErrorIndicatorClickHandler } from '../../../src/mount/error-indicator-click.js';
import { gridHitState } from '../../../src/render/grid/hit-state.js';

const address = { sheet: 0, row: 5, col: 0 };
const triangle = { addr: address, kind: 'error' as const, rect: { x: 32, y: 117, w: 6, h: 6 } };

function setup() {
  const host = document.createElement('div');
  const grid = document.createElement('div');
  const canvas = document.createElement('canvas');
  host.append(grid);
  grid.append(canvas);
  document.body.append(host);
  vi.spyOn(canvas, 'getBoundingClientRect').mockReturnValue(new DOMRect(10, 100, 200, 200));
  gridHitState(grid).errorTriangles = [triangle];
  const open = vi.fn();
  const handler = createErrorIndicatorClickHandler({
    grid,
    canvas,
    getErrorMenu: () => ({ open, close: vi.fn(), detach: vi.fn() }),
  });
  host.addEventListener('pointerdown', handler.onPointerDown);
  host.addEventListener('pointermove', handler.onPointerMove);
  host.addEventListener('pointercancel', handler.onPointerCancel);
  host.addEventListener('click', handler);
  return { host, grid, canvas, handler, open };
}

function pointerDown(canvas: HTMLElement, pointerId = 7, clientX = 45, clientY = 220) {
  canvas.dispatchEvent(
    new PointerEvent('pointerdown', {
      bubbles: true,
      cancelable: true,
      button: 0,
      clientX,
      clientY,
      pointerId,
      isPrimary: true,
    }),
  );
}

function capturedClick(grid: HTMLElement, pointerId = 7, clientX = 45, clientY = 220) {
  grid.dispatchEvent(
    new PointerEvent('click', {
      bubbles: true,
      cancelable: true,
      button: 0,
      clientX,
      clientY,
      pointerId,
    }),
  );
}

afterEach(() => {
  document.body.replaceChildren();
});

describe('error indicator click routing', () => {
  it('opens for a direct canvas click without pointer capture', () => {
    const { canvas, open } = setup();

    canvas.dispatchEvent(
      new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        button: 0,
        clientX: 45,
        clientY: 220,
      }),
    );

    expect(open).toHaveBeenCalledExactlyOnceWith(address, 45, 220, 'error');
  });

  it('opens for a captured grid click armed by a triangle pointerdown', () => {
    const { canvas, grid, open } = setup();

    pointerDown(canvas);
    capturedClick(grid);

    expect(open).toHaveBeenCalledExactlyOnceWith(address, 45, 220, 'error');
  });

  it('does not open for a grid click without a triangle pointerdown', () => {
    const { grid, open } = setup();

    capturedClick(grid);

    expect(open).not.toHaveBeenCalled();
  });

  it('ignores a secondary pointer and a click from a different pointer', () => {
    const { canvas, grid, open } = setup();
    canvas.dispatchEvent(
      new PointerEvent('pointerdown', {
        bubbles: true,
        button: 0,
        pointerId: 8,
        isPrimary: false,
        clientX: 45,
        clientY: 220,
      }),
    );
    capturedClick(grid, 8);
    expect(open).not.toHaveBeenCalled();
    pointerDown(canvas, 7);
    capturedClick(grid, 8);
    expect(open).not.toHaveBeenCalled();
  });

  it('ignores overlay controls at the same screen coordinates as a triangle', () => {
    const { grid, open } = setup();
    const button = document.createElement('button');
    grid.append(button);

    button.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 45, clientY: 220 }));

    expect(open).not.toHaveBeenCalled();
  });

  it('disarms a captured click when the pointer leaves and returns to the triangle', () => {
    const { canvas, grid, open } = setup();

    pointerDown(canvas);
    grid.dispatchEvent(
      new PointerEvent('pointermove', {
        bubbles: true,
        clientX: 100,
        clientY: 300,
        pointerId: 7,
      }),
    );
    grid.dispatchEvent(
      new PointerEvent('pointermove', {
        bubbles: true,
        clientX: 45,
        clientY: 220,
        pointerId: 7,
      }),
    );
    capturedClick(grid);

    expect(open).not.toHaveBeenCalled();
  });

  it('clears an armed click on pointercancel and on feature reset', () => {
    const first = setup();
    pointerDown(first.canvas);
    first.grid.dispatchEvent(new PointerEvent('pointercancel', { bubbles: true, pointerId: 7 }));
    capturedClick(first.grid);
    expect(first.open).not.toHaveBeenCalled();

    const second = setup();
    pointerDown(second.canvas);
    second.handler.reset();
    capturedClick(second.grid);
    expect(second.open).not.toHaveBeenCalled();
  });

  it('requires the captured click to resolve to the same triangle', () => {
    const { canvas, grid, open } = setup();
    pointerDown(canvas);
    gridHitState(grid).errorTriangles = [
      {
        addr: { sheet: 0, row: 6, col: 0 },
        kind: 'error',
        rect: { x: 80, y: 117, w: 6, h: 6 },
      },
    ];
    capturedClick(grid, 7, 93, 220);

    expect(open).not.toHaveBeenCalled();
  });
});
