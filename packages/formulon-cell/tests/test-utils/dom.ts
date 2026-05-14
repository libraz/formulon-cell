import { vi } from 'vitest';

/** A happy-dom canvas Proxy that no-ops draw calls and returns a usable
 *  `measureText`. Mirrors the inline pattern used in
 *  `mount-features.test.ts`. Use via {@link installCanvasStub}. */
function makeCanvasContext(): CanvasRenderingContext2D {
  return new Proxy(
    {
      canvas: document.createElement('canvas'),
      measureText: (text: string) => ({ width: text.length * 7 }),
    },
    {
      get(target, prop) {
        if (prop in target) return target[prop as keyof typeof target];
        return vi.fn();
      },
      set(target, prop, value) {
        (target as Record<PropertyKey, unknown>)[prop] = value;
        return true;
      },
    },
  ) as unknown as CanvasRenderingContext2D;
}

class StubResizeObserver implements ResizeObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
}

/** Install happy-dom-safe `HTMLCanvasElement.getContext` and `ResizeObserver`
 *  stubs that let `Spreadsheet.mount` complete without a real canvas. Pair
 *  with {@link uninstallDomStubs} in `afterEach`. */
export function installDomStubs(): void {
  vi.stubGlobal('ResizeObserver', StubResizeObserver);
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(makeCanvasContext());
}

export function uninstallDomStubs(): void {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
}

/** Create a host div attached to `document.body`, suitable for `Spreadsheet.mount`.
 *  Returns a `cleanup` function that detaches it. */
export function createHostElement(): { host: HTMLElement; cleanup: () => void } {
  const host = document.createElement('div');
  document.body.appendChild(host);
  return {
    host,
    cleanup: () => {
      host.remove();
    },
  };
}

/** Drain all currently-queued microtasks. Useful when `Spreadsheet.mount` uses
 *  await chains that resolve synchronously. */
export async function flushMicrotasks(): Promise<void> {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
}
