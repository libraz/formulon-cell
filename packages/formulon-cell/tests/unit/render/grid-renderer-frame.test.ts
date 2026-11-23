import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { GridRenderer } from '../../../src/render/grid.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

const theme = {
  accent: '#0078d4',
  accentSoft: 'rgba(0,120,212,0.1)',
  bg: '#ffffff',
  bgElev: '#f5f5f5',
  bgRail: '#eeeeee',
  fg: '#1f1f1f',
  cellAutoFg: '#000000',
  rule: '#cccccc',
  fontUi: 'Arial',
  fontMono: 'monospace',
  textCell: 12,
} as unknown as ResolvedTheme;

/** A context that accepts any call or property write. */
function looseCtx(): CanvasRenderingContext2D {
  const store: Record<string | symbol, unknown> = {};
  return new Proxy(store, {
    get(target, key) {
      if (key in target) return target[key];
      if (key === 'measureText') return () => ({ width: 0 });
      return () => ({ addColorStop() {} });
    },
    set(target, key, value) {
      target[key] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

interface Harness {
  renderer: GridRenderer;
  canvas: { width: number; height: number; style: Record<string, string> };
  host: { closest: ReturnType<typeof vi.fn> };
  viewportCalls: () => number;
  sizeWrites: () => number;
}

function harness(withViewport = true, onSize?: () => void): Harness {
  const state = createSpreadsheetStore().getState();
  let viewportCalls = 0;
  let sizeWrites = 0;
  let width = 0;
  let height = 0;
  const canvas = {
    get width() {
      return width;
    },
    set width(v: number) {
      sizeWrites += 1;
      width = v;
    },
    get height() {
      return height;
    },
    set height(v: number) {
      sizeWrites += 1;
      height = v;
    },
    style: {} as Record<string, string>,
    getContext: () => looseCtx(),
  };
  const host = {
    getBoundingClientRect: () => ({ width: 640, height: 480 }),
    closest: vi.fn(() => ({ dataset: { fcPlatform: 'mac' } })),
  };
  const renderer = new GridRenderer({
    host: host as unknown as HTMLElement,
    canvas: canvas as unknown as HTMLCanvasElement,
    getState: () => state,
    getTheme: () => theme,
    onViewportSize: withViewport
      ? () => {
          viewportCalls += 1;
          onSize?.();
        }
      : undefined,
  });
  return {
    renderer,
    canvas,
    host,
    viewportCalls: () => viewportCalls,
    sizeWrites: () => sizeWrites,
  };
}

describe('GridRenderer frame work', () => {
  let rafQueue: FrameRequestCallback[];

  beforeEach(() => {
    rafQueue = [];
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => rafQueue.push(cb));
    vi.stubGlobal('cancelAnimationFrame', () => {});
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('runs the viewport sizing once for repeated invalidates within a frame', () => {
    const h = harness();
    h.renderer.resize();
    const before = h.viewportCalls();
    h.renderer.invalidate();
    h.renderer.invalidate();
    h.renderer.invalidate();
    expect(h.viewportCalls() - before).toBeLessThanOrEqual(1);
    h.renderer.dispose();
  });

  it('keeps one frame pending when the viewport callback invalidates re-entrantly', () => {
    const cancelled = new Set<number>();
    vi.stubGlobal('cancelAnimationFrame', (id: number) => cancelled.add(id));
    let reenter = false;
    const h = harness(true, () => {
      if (!reenter) return;
      reenter = false;
      h.renderer.invalidate();
    });
    h.renderer.resize();
    for (const cb of rafQueue.splice(0)) cb(0);
    reenter = true;
    h.renderer.invalidate();
    expect(rafQueue).toHaveLength(1);
    h.renderer.dispose();
    expect([...cancelled]).toEqual([1]);
  });

  it('resolves the platform profile once per paint', () => {
    const h = harness(false);
    h.renderer.resize();
    h.host.closest.mockClear();
    (h.renderer as unknown as { paint(): void }).paint();
    expect(h.host.closest.mock.calls.length).toBeLessThanOrEqual(1);
    h.renderer.dispose();
  });
});

describe('GridRenderer device pixel ratio', () => {
  type Listener = () => void;
  let queries: { query: string; listeners: Set<Listener> }[];

  beforeEach(() => {
    queries = [];
    vi.stubGlobal('requestAnimationFrame', () => 1);
    vi.stubGlobal('cancelAnimationFrame', () => {});
    vi.stubGlobal('devicePixelRatio', 1);
    vi.stubGlobal('matchMedia', (query: string) => {
      const entry = { query, listeners: new Set<Listener>() };
      queries.push(entry);
      return {
        media: query,
        matches: true,
        addEventListener: (_: string, l: Listener) => entry.listeners.add(l),
        removeEventListener: (_: string, l: Listener) => entry.listeners.delete(l),
      };
    });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('rescales the backing store when the display resolution changes', () => {
    const h = harness(false);
    h.renderer.resize();
    expect(h.canvas.width).toBe(640);
    const first = queries.at(-1);
    expect(first?.query).toBe('(resolution: 1dppx)');

    vi.stubGlobal('devicePixelRatio', 2);
    for (const l of [...(first?.listeners ?? [])]) l();
    expect(h.canvas.width).toBe(1280);
    expect(queries.at(-1)?.query).toBe('(resolution: 2dppx)');
    h.renderer.dispose();
  });

  it('stops listening after dispose', () => {
    const h = harness(false);
    h.renderer.resize();
    h.renderer.dispose();
    expect(queries.every((q) => q.listeners.size === 0)).toBe(true);
  });

  it('leaves the backing store alone when the size is unchanged', () => {
    const h = harness(false);
    h.renderer.resize();
    const writes = h.sizeWrites();
    h.renderer.resize();
    expect(h.sizeWrites()).toBe(writes);
    h.renderer.dispose();
  });
});
