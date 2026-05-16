import type { SpreadsheetInstance } from '@libraz/formulon-cell';

import { SVG_NS } from './ribbon/border-icons.js';

export type SessionIllustrationKind = 'image' | 'shape' | 'screenshot';
export type SessionShapeKind = 'rectangle' | 'rounded-rectangle' | 'oval' | 'line' | 'arrow';

export type SessionIllustration = {
  id: string;
  kind: SessionIllustrationKind;
  shape?: SessionShapeKind;
  url?: string;
  x: number;
  y: number;
  w: number;
  h: number;
};

// Minimal subset of `ribbonMenuText` consumed for accessible labels.
export type IllustrationLabels = {
  pictureOnline: string;
  screenshotCurrentView: string;
  shapeRoundedRectangle: string;
  shapeOval: string;
  shapeLine: string;
  shapeArrow: string;
  shapeRectangle: string;
};

export interface IllustrationsCtx {
  getInst: () => SpreadsheetInstance | null;
  getSheetEl: () => HTMLElement | null;
  getLabels: () => IllustrationLabels;
  focusSheet: () => void;
}

export interface IllustrationsApi {
  addSessionIllustration: (
    kind: SessionIllustrationKind,
    input?: Partial<SessionIllustration>,
  ) => void;
  setDrawInkMode: (mode: 'pen' | 'erase') => void;
}

export const createIllustrations = (ctx: IllustrationsCtx): IllustrationsApi => {
  const sessionIllustrations: SessionIllustration[] = [];
  let selectedIllustrationId: string | null = null;

  const cloneSessionIllustration = (item: SessionIllustration): SessionIllustration => ({
    ...item,
  });

  const captureSessionIllustrationsSnapshot = (): {
    items: SessionIllustration[];
    selectedId: string | null;
  } => ({
    items: sessionIllustrations.map(cloneSessionIllustration),
    selectedId: selectedIllustrationId,
  });

  const applySessionIllustrationsSnapshot = (snapshot: {
    items: readonly SessionIllustration[];
    selectedId: string | null;
  }): void => {
    sessionIllustrations.splice(
      0,
      sessionIllustrations.length,
      ...snapshot.items.map(cloneSessionIllustration),
    );
    selectedIllustrationId = snapshot.selectedId;
    renderSessionIllustrations();
  };

  const sameSessionIllustrationsSnapshot = (
    a: { items: readonly SessionIllustration[]; selectedId: string | null },
    b: { items: readonly SessionIllustration[]; selectedId: string | null },
  ): boolean =>
    a.selectedId === b.selectedId && JSON.stringify(a.items) === JSON.stringify(b.items);

  const recordSessionIllustrationsChange = (mutate: () => void): void => {
    const history = ctx.getInst()?.history ?? null;
    if (!history || history.isReplaying()) {
      mutate();
      return;
    }
    const before = captureSessionIllustrationsSnapshot();
    mutate();
    const after = captureSessionIllustrationsSnapshot();
    if (sameSessionIllustrationsSnapshot(before, after)) return;
    history.push({
      undo: () => applySessionIllustrationsSnapshot(before),
      redo: () => applySessionIllustrationsSnapshot(after),
    });
  };

  type SessionInkStroke = {
    id: string;
    points: Array<{ x: number; y: number }>;
  };

  const sessionInkStrokes: SessionInkStroke[] = [];
  let drawInkMode: 'pen' | 'erase' | null = null;
  let inkListenersAttached = false;

  const cloneSessionInkStroke = (stroke: SessionInkStroke): SessionInkStroke => ({
    id: stroke.id,
    points: stroke.points.map((point) => ({ ...point })),
  });

  const captureSessionInkSnapshot = (): SessionInkStroke[] =>
    sessionInkStrokes.map(cloneSessionInkStroke);

  const applySessionInkSnapshot = (snapshot: readonly SessionInkStroke[]): void => {
    sessionInkStrokes.splice(0, sessionInkStrokes.length, ...snapshot.map(cloneSessionInkStroke));
    renderSessionInk();
  };

  const sameSessionInkSnapshot = (
    a: readonly SessionInkStroke[],
    b: readonly SessionInkStroke[],
  ): boolean => JSON.stringify(a) === JSON.stringify(b);

  const pushSessionInkHistory = (
    before: readonly SessionInkStroke[],
    after: readonly SessionInkStroke[],
  ): void => {
    const history = ctx.getInst()?.history ?? null;
    if (!history || history.isReplaying() || sameSessionInkSnapshot(before, after)) return;
    const undoSnapshot = before.map(cloneSessionInkStroke);
    const redoSnapshot = after.map(cloneSessionInkStroke);
    history.push({
      undo: () => applySessionInkSnapshot(undoSnapshot),
      redo: () => applySessionInkSnapshot(redoSnapshot),
    });
  };

  const recordSessionInkChange = (mutate: () => void): void => {
    const before = captureSessionInkSnapshot();
    mutate();
    pushSessionInkHistory(before, captureSessionInkSnapshot());
  };

  const illustrationGrid = (): HTMLElement | null =>
    ctx.getSheetEl()?.querySelector<HTMLElement>('.fc-host__grid') ?? null;

  const syncDrawInkButtons = (): void => {
    for (const button of document.querySelectorAll<HTMLButtonElement>(
      '[data-ribbon-command="drawPen"], [data-ribbon-command="drawErase"]',
    )) {
      const active =
        (button.dataset.ribbonCommand === 'drawPen' && drawInkMode === 'pen') ||
        (button.dataset.ribbonCommand === 'drawErase' && drawInkMode === 'erase');
      button.setAttribute('aria-pressed', active ? 'true' : 'false');
    }
  };

  const inkRoot = (): SVGSVGElement | null => {
    const grid = illustrationGrid();
    if (!grid) return null;
    let root = grid.querySelector<SVGSVGElement>('.app-ink');
    if (!root) {
      root = document.createElementNS(SVG_NS, 'svg');
      root.classList.add('app-ink');
      root.setAttribute('aria-hidden', 'true');
      grid.appendChild(root);
    }
    return root;
  };

  const renderSessionInk = (): void => {
    const root = inkRoot();
    if (!root) return;
    root.replaceChildren();
    for (const stroke of sessionInkStrokes) {
      const polyline = document.createElementNS(SVG_NS, 'polyline');
      polyline.classList.add('app-ink__stroke');
      polyline.dataset.inkStrokeId = stroke.id;
      polyline.setAttribute('points', stroke.points.map((p) => `${p.x},${p.y}`).join(' '));
      root.appendChild(polyline);
    }
  };

  const gridPointFromPointer = (event: PointerEvent): { x: number; y: number } | null => {
    const grid = illustrationGrid();
    if (!grid) return null;
    const rect = grid.getBoundingClientRect();
    return {
      x: Math.max(0, event.clientX - rect.left),
      y: Math.max(0, event.clientY - rect.top),
    };
  };

  const pointToSegmentDistance = (
    point: { x: number; y: number },
    a: { x: number; y: number },
    b: { x: number; y: number },
  ): number => {
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const len2 = dx * dx + dy * dy || 1;
    const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / len2));
    const x = a.x + t * dx;
    const y = a.y + t * dy;
    return Math.hypot(point.x - x, point.y - y);
  };

  const eraseInkAt = (point: { x: number; y: number }): void => {
    const index = sessionInkStrokes.findIndex((stroke) => {
      if (stroke.points.length === 1) {
        return Math.hypot(stroke.points[0]!.x - point.x, stroke.points[0]!.y - point.y) < 12;
      }
      for (let i = 1; i < stroke.points.length; i += 1) {
        if (pointToSegmentDistance(point, stroke.points[i - 1]!, stroke.points[i]!) < 12) {
          return true;
        }
      }
      return false;
    });
    if (index >= 0) {
      recordSessionInkChange(() => {
        sessionInkStrokes.splice(index, 1);
        renderSessionInk();
      });
    }
  };

  const attachInkPointerListeners = (): void => {
    const grid = illustrationGrid();
    if (!grid || inkListenersAttached) return;
    inkListenersAttached = true;
    grid.addEventListener('pointerdown', (event) => {
      if (!drawInkMode || event.button !== 0) return;
      const point = gridPointFromPointer(event);
      if (!point) return;
      event.preventDefault();
      event.stopPropagation();
      if (drawInkMode === 'erase') {
        eraseInkAt(point);
        return;
      }
      const before = captureSessionInkSnapshot();
      const stroke: SessionInkStroke = {
        id: `ink-${Date.now().toString(36)}-${sessionInkStrokes.length}`,
        points: [point],
      };
      sessionInkStrokes.push(stroke);
      renderSessionInk();
      const onMove = (moveEvent: PointerEvent): void => {
        const next = gridPointFromPointer(moveEvent);
        if (!next) return;
        stroke.points.push(next);
        renderSessionInk();
      };
      const onUp = (): void => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        window.removeEventListener('pointercancel', onUp);
        renderSessionInk();
        pushSessionInkHistory(before, captureSessionInkSnapshot());
      };
      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
      window.addEventListener('pointercancel', onUp);
    });
  };

  const setDrawInkMode = (mode: 'pen' | 'erase'): void => {
    drawInkMode = drawInkMode === mode ? null : mode;
    attachInkPointerListeners();
    inkRoot();
    const sheetEl = ctx.getSheetEl();
    sheetEl
      ?.querySelector<HTMLElement>('.fc-host')
      ?.classList.toggle('app-ink--pen', drawInkMode === 'pen');
    sheetEl
      ?.querySelector<HTMLElement>('.fc-host')
      ?.classList.toggle('app-ink--erase', drawInkMode === 'erase');
    syncDrawInkButtons();
    illustrationGrid()?.focus();
  };

  const illustrationRoot = (): HTMLElement | null => {
    const grid = illustrationGrid();
    if (!grid) return null;
    let root = grid.querySelector<HTMLElement>('.app-illustrations');
    if (!root) {
      root = document.createElement('div');
      root.className = 'app-illustrations';
      grid.appendChild(root);
    }
    return root;
  };

  const updateSessionIllustration = (
    id: string,
    patch: Partial<Pick<SessionIllustration, 'x' | 'y' | 'w' | 'h'>>,
  ): void => {
    recordSessionIllustrationsChange(() => {
      const item = sessionIllustrations.find((candidate) => candidate.id === id);
      if (!item) return;
      Object.assign(item, patch);
      renderSessionIllustrations();
    });
  };

  const removeSessionIllustration = (id: string): void => {
    recordSessionIllustrationsChange(() => {
      const index = sessionIllustrations.findIndex((candidate) => candidate.id === id);
      if (index < 0) return;
      sessionIllustrations.splice(index, 1);
      if (selectedIllustrationId === id) selectedIllustrationId = null;
      renderSessionIllustrations();
    });
  };

  const applyIllustrationPointerDrag = (
    event: PointerEvent,
    item: SessionIllustration,
    node: HTMLElement,
    mode: 'move' | 'resize',
  ): void => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    selectedIllustrationId = item.id;
    illustrationRoot()
      ?.querySelectorAll<HTMLElement>('.app-illustration')
      .forEach((candidate) =>
        candidate.setAttribute(
          'aria-selected',
          candidate.dataset.illustrationId === item.id ? 'true' : 'false',
        ),
      );
    node.focus();
    const startX = event.clientX;
    const startY = event.clientY;
    const start = { x: item.x, y: item.y, w: item.w, h: item.h };
    let next = { ...start };
    const applyLive = (): void => {
      node.style.left = `${next.x}px`;
      node.style.top = `${next.y}px`;
      node.style.width = `${next.w}px`;
      node.style.height = `${next.h}px`;
    };
    const onMove = (moveEvent: PointerEvent): void => {
      const dx = moveEvent.clientX - startX;
      const dy = moveEvent.clientY - startY;
      if (mode === 'resize') {
        next = {
          ...start,
          w: Math.max(24, start.w + dx),
          h: Math.max(12, start.h + dy),
        };
      } else {
        next = {
          ...start,
          x: Math.max(0, start.x + dx),
          y: Math.max(0, start.y + dy),
        };
      }
      applyLive();
    };
    const onUp = (upEvent: PointerEvent): void => {
      upEvent.preventDefault();
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onUp);
      updateSessionIllustration(item.id, next);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onUp);
  };

  const applyIllustrationKeyboard = (
    event: KeyboardEvent,
    item: SessionIllustration,
    node: HTMLElement,
  ): void => {
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      removeSessionIllustration(item.id);
      ctx.focusSheet();
      return;
    }
    const step = event.shiftKey ? 10 : 1;
    const resize = event.altKey;
    const delta: [number, number] | null =
      event.key === 'ArrowLeft'
        ? [-step, 0]
        : event.key === 'ArrowRight'
          ? [step, 0]
          : event.key === 'ArrowUp'
            ? [0, -step]
            : event.key === 'ArrowDown'
              ? [0, step]
              : null;
    if (!delta) return;
    event.preventDefault();
    if (resize) {
      updateSessionIllustration(item.id, {
        w: Math.max(24, item.w + delta[0]),
        h: Math.max(12, item.h + delta[1]),
      });
    } else {
      updateSessionIllustration(item.id, {
        x: Math.max(0, item.x + delta[0]),
        y: Math.max(0, item.y + delta[1]),
      });
    }
    node.focus();
  };

  const illustrationLabel = (item: SessionIllustration): string => {
    const t = ctx.getLabels();
    if (item.kind === 'image') return t.pictureOnline;
    if (item.kind === 'screenshot') return t.screenshotCurrentView;
    if (item.shape === 'rounded-rectangle') return t.shapeRoundedRectangle;
    if (item.shape === 'oval') return t.shapeOval;
    if (item.shape === 'line') return t.shapeLine;
    if (item.shape === 'arrow') return t.shapeArrow;
    return t.shapeRectangle;
  };

  const renderSessionIllustrations = (): void => {
    const root = illustrationRoot();
    if (!root) return;
    root.replaceChildren();
    for (const item of sessionIllustrations) {
      const node = document.createElement('div');
      node.className = `app-illustration app-illustration--${item.kind}`;
      node.setAttribute('role', 'button');
      node.tabIndex = 0;
      if (item.shape) node.classList.add(`app-illustration--${item.shape}`);
      node.dataset.illustrationId = item.id;
      node.dataset.illustrationType = item.kind;
      if (item.shape) node.dataset.shape = item.shape;
      node.setAttribute('aria-label', illustrationLabel(item));
      node.setAttribute('aria-selected', item.id === selectedIllustrationId ? 'true' : 'false');
      node.style.left = `${item.x}px`;
      node.style.top = `${item.y}px`;
      node.style.width = `${item.w}px`;
      node.style.height = `${item.h}px`;
      node.addEventListener('pointerdown', (event) => {
        const rect = node.getBoundingClientRect();
        const nearResizeHandle =
          event.clientX >= rect.right - 14 && event.clientY >= rect.bottom - 14;
        applyIllustrationPointerDrag(event, item, node, nearResizeHandle ? 'resize' : 'move');
      });
      node.addEventListener('keydown', (event) => applyIllustrationKeyboard(event, item, node));
      if (item.kind === 'image') {
        const image = document.createElement('img');
        image.alt = '';
        image.src = item.url ?? '';
        node.appendChild(image);
      } else if (item.kind === 'screenshot') {
        node.appendChild(document.createElement('span'));
      }
      const resize = document.createElement('span');
      resize.className = 'app-illustration__resize';
      resize.setAttribute('aria-hidden', 'true');
      resize.addEventListener('pointerdown', (event) =>
        applyIllustrationPointerDrag(event, item, node, 'resize'),
      );
      node.appendChild(resize);
      root.appendChild(node);
    }
  };

  const addSessionIllustration = (
    kind: SessionIllustrationKind,
    input: Partial<SessionIllustration> = {},
  ): void => {
    const count = sessionIllustrations.length;
    const item: SessionIllustration = {
      id: `illustration-${Date.now().toString(36)}-${count}`,
      kind,
      x: 360 + (count % 4) * 28,
      y: 340 + (count % 4) * 24,
      w: kind === 'shape' && (input.shape === 'line' || input.shape === 'arrow') ? 150 : 180,
      h: kind === 'shape' && (input.shape === 'line' || input.shape === 'arrow') ? 28 : 110,
      ...input,
    };
    recordSessionIllustrationsChange(() => {
      sessionIllustrations.push(item);
      selectedIllustrationId = item.id;
      renderSessionIllustrations();
    });
    illustrationGrid()?.focus();
  };

  return {
    addSessionIllustration,
    setDrawInkMode,
  };
};
