import { vi } from 'vitest';

// happy-dom has no 2D canvas and no ResizeObserver; the grid needs both to mount. The stand-ins draw
// nothing and measure text at a fixed advance, which is all a consumer-side test reads.
const canvasContext = (): CanvasRenderingContext2D =>
  new Proxy(
    {
      canvas: document.createElement('canvas'),
      measureText: (text: string) => ({ width: text.length * 7 }),
      createLinearGradient: () => ({ addColorStop: (): undefined => undefined }),
      createRadialGradient: () => ({ addColorStop: (): undefined => undefined }),
    },
    {
      get: (target, prop) =>
        prop in target ? target[prop as keyof typeof target] : (): undefined => undefined,
      set: (target, prop, value) => {
        (target as Record<PropertyKey, unknown>)[prop] = value;
        return true;
      },
    },
  ) as unknown as CanvasRenderingContext2D;

HTMLCanvasElement.prototype.getContext = function getContext() {
  return canvasContext();
} as unknown as typeof HTMLCanvasElement.prototype.getContext;

if (typeof globalThis.ResizeObserver === 'undefined') {
  vi.stubGlobal(
    'ResizeObserver',
    class {
      observe(): void {}
      unobserve(): void {}
      disconnect(): void {}
    },
  );
}
