import { mount, type SpreadsheetInstance, Workbook } from '@libraz/formulon-cell';
import { vi } from 'vitest';
import { nextTick } from 'vue';

// happy-dom has no 2D canvas and no ResizeObserver; the grid needs both to mount. The stand-ins
// draw nothing and measure text at a fixed advance.
const canvasContext = (): CanvasRenderingContext2D =>
  new Proxy(
    {
      canvas: document.createElement('canvas'),
      measureText: (text: string) => ({ width: text.length * 7 }),
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

/** Lets pending promise continuations and Vue's scheduler run. */
export const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
  await nextTick();
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

/** Polls until `predicate` holds, flushing between attempts. */
export const waitFor = async (predicate: () => boolean, attempts = 100): Promise<void> => {
  for (let i = 0; i < attempts && !predicate(); i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 5));
    await flush();
  }
  if (!predicate()) throw new Error('waitFor: condition never became true');
};

/** A detached host element of a fixed size, appended to the document. */
export const makeHost = (): HTMLElement => {
  const host = document.createElement('div');
  host.style.width = '800px';
  host.style.height = '600px';
  document.body.appendChild(host);
  return host;
};

/** A workbook on the in-memory stub engine. */
export const stubWorkbook = (): Promise<Workbook> => Workbook.createDefault({ engine: 'stub' });

/** Mounts a core instance on a stub workbook the instance does not own. */
export const mountStub = async (
  options: Parameters<typeof mount>[1] = {},
): Promise<SpreadsheetInstance> =>
  mount(makeHost(), { locale: 'en', workbook: await stubWorkbook(), ...options });
