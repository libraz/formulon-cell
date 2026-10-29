import type { ErrorMenuHandle } from '../interact/error-menu.js';
import type { ErrorTriangleHit } from '../render/grid/hit-state.js';
import { gridHitState } from '../render/grid.js';

export interface ErrorIndicatorClickDeps {
  grid: HTMLElement;
  canvas: HTMLElement;
  getErrorMenu: () => ErrorMenuHandle | null;
}

export type ErrorIndicatorClickHandler = ((e: MouseEvent) => void) & {
  onPointerDown: (e: PointerEvent) => void;
  onPointerMove: (e: PointerEvent) => void;
  onPointerCancel: (e: PointerEvent) => void;
  reset: () => void;
};

interface ArmedTriangle {
  pointerId: number;
  hit: ErrorTriangleHit;
}

const HIT_PADDING = 2;

function pointerIdOf(e: MouseEvent): number | null {
  const pointerId = (e as MouseEvent & { pointerId?: unknown }).pointerId;
  return typeof pointerId === 'number' && Number.isFinite(pointerId) ? pointerId : null;
}

function sameHit(a: ErrorTriangleHit, b: ErrorTriangleHit): boolean {
  return (
    a.kind === b.kind &&
    a.addr.sheet === b.addr.sheet &&
    a.addr.row === b.addr.row &&
    a.addr.col === b.addr.col
  );
}

function hitContains(hit: ErrorTriangleHit, x: number, y: number): boolean {
  const r = hit.rect;
  return (
    x >= r.x - HIT_PADDING &&
    x <= r.x + r.w + HIT_PADDING &&
    y >= r.y - HIT_PADDING &&
    y <= r.y + r.h + HIT_PADDING
  );
}

// Error / validation triangle clicks. We use `click` (not `pointerdown`)
// so the existing pointer handler gets to set the active cell first —
// that way the menu and the cell select agree on the addr the user
// just clicked.
export function createErrorIndicatorClickHandler(
  deps: ErrorIndicatorClickDeps,
): ErrorIndicatorClickHandler {
  const { grid, canvas, getErrorMenu } = deps;
  let armed: ArmedTriangle | null = null;

  const localPoint = (e: MouseEvent): { x: number; y: number } => {
    const rect = canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  const hitAt = (e: MouseEvent): ErrorTriangleHit | null => {
    const { x, y } = localPoint(e);
    return gridHitState(grid).errorTriangles.find((hit) => hitContains(hit, x, y)) ?? null;
  };

  const reset = (): void => {
    armed = null;
  };

  const onPointerDown = (e: PointerEvent): void => {
    // Every new gesture replaces the previous click intent. Only a primary
    // pointer that really started on the painted canvas can arm a captured
    // grid click; overlay controls and stale captures therefore cannot open a
    // menu underneath them.
    reset();
    if (e.button !== 0 || e.isPrimary === false || e.target !== canvas) return;
    const hit = hitAt(e);
    if (hit) armed = { pointerId: e.pointerId, hit };
  };

  const onPointerMove = (e: PointerEvent): void => {
    if (!armed || e.pointerId !== armed.pointerId) return;
    const { x, y } = localPoint(e);
    if (!hitContains(armed.hit, x, y)) reset();
  };

  const onPointerCancel = (e: PointerEvent): void => {
    if (!armed || e.pointerId === armed.pointerId) reset();
  };

  const onClick = ((e: MouseEvent): void => {
    // Consume the intent before any early return. A click is the end of one
    // gesture, even when it lands on an overlay or uses a non-primary button.
    const captured = armed;
    reset();

    if (e.button !== 0) return;
    // Pointer capture sends a canvas click to the grid. Child controls must
    // not activate a triangle underneath their overlay.
    if (e.target !== canvas && e.target !== grid) return;
    const hit = hitAt(e);
    if (!hit) return;
    if (e.target === grid) {
      const pointerId = pointerIdOf(e);
      if (!captured || (pointerId !== null && pointerId !== captured.pointerId)) return;
      if (!sameHit(hit, captured.hit)) return;
    }
    const errorMenu = getErrorMenu();
    if (!errorMenu) return;
    e.stopPropagation();
    e.preventDefault();
    errorMenu.open(hit.addr, e.clientX, e.clientY, hit.kind);
  }) as ErrorIndicatorClickHandler;

  onClick.onPointerDown = onPointerDown;
  onClick.onPointerMove = onPointerMove;
  onClick.onPointerCancel = onPointerCancel;
  onClick.reset = reset;
  return onClick;
}
