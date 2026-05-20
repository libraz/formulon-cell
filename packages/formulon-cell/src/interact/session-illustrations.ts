import type { History } from '../commands/history.js';
import {
  clearSessionIllustration,
  updateSessionIllustration,
} from '../commands/session-illustration.js';
import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface SessionIllustrationsHandle {
  refresh(): void;
  detach(): void;
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const MIN_W = 48;
const MIN_H = 32;
const KEYBOARD_STEP = 8;

const appendSvg = <K extends keyof SVGElementTagNameMap>(
  parent: Element,
  tag: K,
  attrs: Record<string, string | number>,
): SVGElementTagNameMap[K] => {
  const el = document.createElementNS(SVG_NS, tag);
  for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, String(value));
  parent.appendChild(el);
  return el;
};

const renderShape = (svg: SVGSVGElement, item: SessionIllustration): void => {
  svg.replaceChildren();
  const color = item.color ?? '#0f6cbd';
  if (!item.shape) return;
  if (item.shape === 'line' || item.shape === 'arrow') {
    appendSvg(svg, 'line', {
      x1: 12,
      y1: 68,
      x2: 148,
      y2: 20,
      stroke: color,
      'stroke-width': 5,
      'stroke-linecap': 'round',
    });
    if (item.shape === 'arrow') {
      appendSvg(svg, 'path', {
        d: 'M 148 20 L 132 18 L 142 34 Z',
        fill: color,
      });
    }
    return;
  }
  if (item.shape === 'oval') {
    appendSvg(svg, 'ellipse', {
      cx: 80,
      cy: 48,
      rx: 68,
      ry: 36,
      fill: color,
      'fill-opacity': 0.16,
      stroke: color,
      'stroke-width': 3,
    });
    return;
  }
  if (item.shape === 'triangle') {
    appendSvg(svg, 'polygon', {
      points: '80,12 148,84 12,84',
      fill: color,
      'fill-opacity': 0.16,
      stroke: color,
      'stroke-width': 3,
      'stroke-linejoin': 'round',
    });
    return;
  }
  if (item.shape === 'diamond') {
    appendSvg(svg, 'polygon', {
      points: '80,10 150,48 80,86 10,48',
      fill: color,
      'fill-opacity': 0.16,
      stroke: color,
      'stroke-width': 3,
      'stroke-linejoin': 'round',
    });
    return;
  }
  appendSvg(svg, 'rect', {
    x: 12,
    y: 12,
    width: 136,
    height: 72,
    rx: item.shape === 'rounded-rectangle' ? 12 : 2,
    fill: color,
    'fill-opacity': 0.16,
    stroke: color,
    'stroke-width': 3,
  });
};

const appendImage = (
  panel: HTMLElement,
  item: SessionIllustration,
  onPointerDown: (event: PointerEvent) => void,
): void => {
  const img = document.createElement('img');
  img.src = item.src ?? '';
  img.alt = item.alt ?? '';
  img.draggable = false;
  img.style.cssText =
    'display:block;width:100%;height:100%;object-fit:contain;pointer-events:auto;';
  img.addEventListener('pointerdown', onPointerDown);
  panel.appendChild(img);
};

export function attachSessionIllustrations(deps: {
  host: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
  closeLabel?: string;
  pictureLabel?: string;
  shapeLabel?: string;
  resizeLabel?: string;
}): SessionIllustrationsHandle {
  const { host, store, history = null } = deps;
  const root = document.createElement('div');
  root.className = 'fc-illustrations';
  root.style.cssText = 'position:absolute;inset:0;pointer-events:none;';
  host.appendChild(root);
  inheritHostTokens(host, root);
  let selectedId: string | null = null;

  const select = (id: string, panel: HTMLElement): void => {
    selectedId = id;
    refreshSelection();
    if (document.activeElement !== panel) panel.focus();
  };

  const refreshSelection = (): void => {
    const panels = Array.from(root.querySelectorAll<HTMLElement>('.fc-illustration'));
    for (const [index, panel] of panels.entries()) {
      const selected = panel.dataset.illustrationId === selectedId;
      panel.classList.toggle('fc-illustration--selected', selected);
      panel.setAttribute('aria-selected', selected ? 'true' : 'false');
      panel.style.outline = selected ? '2px solid var(--fc-accent, #0f6cbd)' : 'none';
      panel.style.zIndex = selected ? '1000' : String(index + 1);
    }
  };

  const applyDrag = (
    e: PointerEvent,
    item: SessionIllustration,
    mode: 'move' | 'resize',
    panel: HTMLElement,
  ): void => {
    if (e.button !== 0) return;
    e.preventDefault();
    const startX = e.clientX;
    const startY = e.clientY;
    const startLeft = item.x ?? panel.offsetLeft;
    const startTop = item.y ?? panel.offsetTop;
    const startW = item.w ?? panel.offsetWidth;
    const startH = item.h ?? panel.offsetHeight;
    const onMove = (move: PointerEvent): void => {
      const dx = move.clientX - startX;
      const dy = move.clientY - startY;
      if (mode === 'move') {
        updateSessionIllustration(
          store,
          item.id,
          {
            x: Math.max(0, startLeft + dx),
            y: Math.max(0, startTop + dy),
          },
          history,
        );
        return;
      }
      updateSessionIllustration(
        store,
        item.id,
        {
          w: Math.max(MIN_W, startW + dx),
          h: Math.max(MIN_H, startH + dy),
        },
        history,
      );
    };
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp, { once: true });
  };

  const applyKeyboard = (event: KeyboardEvent, item: SessionIllustration, panel: HTMLElement) => {
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      clearSessionIllustration(store, item.id, history);
      host.focus({ preventScroll: true });
      return;
    }
    const deltas: Record<string, [number, number]> = {
      ArrowLeft: [-KEYBOARD_STEP, 0],
      ArrowRight: [KEYBOARD_STEP, 0],
      ArrowUp: [0, -KEYBOARD_STEP],
      ArrowDown: [0, KEYBOARD_STEP],
    };
    const delta = deltas[event.key];
    if (!delta) return;
    event.preventDefault();
    const [dx, dy] = delta;
    if (event.shiftKey) {
      updateSessionIllustration(
        store,
        item.id,
        {
          w: Math.max(MIN_W, (item.w ?? panel.offsetWidth) + dx),
          h: Math.max(MIN_H, (item.h ?? panel.offsetHeight) + dy),
        },
        history,
      );
      return;
    }
    updateSessionIllustration(
      store,
      item.id,
      {
        x: Math.max(0, (item.x ?? panel.offsetLeft) + dx),
        y: Math.max(0, (item.y ?? panel.offsetTop) + dy),
      },
      history,
    );
  };

  const render = (): void => {
    const state = store.getState();
    root.replaceChildren();
    state.illustrations.illustrations
      .filter((item) => item.sheet === state.data.sheetIndex)
      .forEach((item) => {
        const panel = document.createElement('section');
        panel.className = 'fc-illustration';
        panel.dataset.illustrationId = item.id;
        panel.tabIndex = 0;
        panel.setAttribute('role', 'group');
        panel.setAttribute(
          'aria-roledescription',
          item.kind === 'image'
            ? (deps.pictureLabel ?? item.kind)
            : (deps.shapeLabel ?? item.kind),
        );
        panel.setAttribute('aria-selected', item.id === selectedId ? 'true' : 'false');
        panel.setAttribute(
          'aria-label',
          item.kind === 'image'
            ? (item.alt ?? deps.pictureLabel ?? item.id)
            : (item.shape ?? deps.shapeLabel ?? item.id),
        );
        panel.style.cssText =
          'position:absolute;pointer-events:auto;background:transparent;box-sizing:border-box;';
        panel.style.left = `${item.x ?? 300}px`;
        panel.style.top = `${item.y ?? 88}px`;
        panel.style.width = `${item.w ?? 160}px`;
        panel.style.height = `${item.h ?? 96}px`;
        panel.addEventListener('focus', () => {
          selectedId = item.id;
          refreshSelection();
        });
        panel.addEventListener('pointerdown', () => select(item.id, panel));
        panel.addEventListener('keydown', (event) => applyKeyboard(event, item, panel));

        if (item.kind === 'image') {
          appendImage(panel, item, (event) => applyDrag(event, item, 'move', panel));
        } else {
          const svg = document.createElementNS(SVG_NS, 'svg');
          svg.setAttribute('viewBox', '0 0 160 96');
          svg.setAttribute('aria-hidden', 'true');
          svg.style.cssText = 'display:block;width:100%;height:100%;';
          svg.addEventListener('pointerdown', (event) => applyDrag(event, item, 'move', panel));
          renderShape(svg, item);
          panel.appendChild(svg);
        }

        const resize = document.createElement('div');
        resize.setAttribute('role', 'separator');
        resize.setAttribute('aria-label', deps.resizeLabel ?? deps.shapeLabel ?? item.id);
        resize.style.cssText =
          'position:absolute;right:-4px;bottom:-4px;width:8px;height:8px;background:var(--fc-accent, #0f6cbd);cursor:nwse-resize;';
        resize.addEventListener('pointerdown', (event) => applyDrag(event, item, 'resize', panel));
        panel.appendChild(resize);
        root.appendChild(panel);
      });
    refreshSelection();
  };

  const unsubscribe = store.subscribe(render);
  render();

  return {
    refresh: render,
    detach() {
      unsubscribe();
      root.remove();
    },
  };
}
