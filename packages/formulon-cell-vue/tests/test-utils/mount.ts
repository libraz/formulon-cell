import { type SpreadsheetInstance, WorkbookHandle } from '@libraz/formulon-cell';
import { vi } from 'vitest';
import {
  type App,
  type ComponentPublicInstance,
  createApp,
  defineComponent,
  h,
  nextTick,
  type Ref,
  shallowRef,
} from 'vue';
import { Spreadsheet, type SpreadsheetExposed } from '../../src/Spreadsheet';

/** Result of {@link mountVueSpreadsheet}. The caller should always call
 *  `dispose()` so the Vue app is unmounted, the host is removed, and the
 *  canvas / ResizeObserver stubs are restored.
 */
export interface MountedVueSpreadsheet {
  host: HTMLDivElement;
  app: App;
  /** Live core instance after mount completes. Throws if mount failed. */
  instance: SpreadsheetInstance;
  /** Exposed `{ instance: Ref<...> }` object via Vue's `expose()`. */
  exposed: SpreadsheetExposed;
  /** Update reactive props by mutating the props ref and awaiting nextTick. */
  setProp: <K extends string>(key: K, value: unknown) => Promise<void>;
  dispose: () => Promise<void>;
}

class StubResizeObserver implements ResizeObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
}

const makeCanvasContext = (): CanvasRenderingContext2D =>
  new Proxy(
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

export const installVueDomStubs = (): void => {
  vi.stubGlobal('ResizeObserver', StubResizeObserver);
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(makeCanvasContext());
};

export const uninstallVueDomStubs = (): void => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
};

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
  await nextTick();
};

interface VueMountOptions {
  /** Initial props to forward to `<Spreadsheet>`. */
  props?: Record<string, unknown>;
  /** Initial event listeners (e.g. `{ onCellChange: fn }`) — Vue treats these
   *  as standard `on{Event}` props on `defineComponent`. */
  listeners?: Record<string, (...args: unknown[]) => void>;
}

/** Mount the wrapper Vue `<Spreadsheet>` against the real core (stub engine)
 *  and return a handle whose `instance` is guaranteed non-null. */
export async function mountVueSpreadsheet(
  options: VueMountOptions = {},
): Promise<MountedVueSpreadsheet> {
  installVueDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);

  // Pre-create a workbook so we don't depend on a default-load codepath that
  // might race the WASM engine in happy-dom.
  const workbook =
    (options.props?.workbook as WorkbookHandle | undefined) ??
    (await WorkbookHandle.createDefault({ preferStub: true }));

  // Reactive prop bag the test can mutate via setProp(). We use `shallowRef`
  // so Vue doesn't deep-walk the workbook handle (a class instance with a
  // live WASM/stub backing). Deep reactivity wraps the workbook in a Proxy,
  // which then fails identity-based comparisons against the original handle.
  const propsState = shallowRef<Record<string, unknown>>({
    workbook,
    ...(options.props ?? {}),
  }) as Ref<Record<string, unknown>>;

  // We capture the live instance via two channels:
  //  1. The `ready` event — fired exactly once with the instance.
  //  2. A template ref on the child component. Vue's exposed proxy unwraps
  //     refs at access time, so `childRef.value.instance` is the unwrapped
  //     instance, not the ref. We mirror it back into a local shallowRef so
  //     the test can keep reading it after prop changes.
  const instanceFromReady = shallowRef<SpreadsheetInstance | null>(null);
  const childRef = shallowRef<ComponentPublicInstance | null>(null);
  const onReady = vi.fn();

  const Wrapper = defineComponent({
    name: 'TestWrapper',
    setup() {
      return () =>
        h(Spreadsheet, {
          ...propsState.value,
          ...(options.listeners ?? {}),
          onReady: (inst: SpreadsheetInstance) => {
            instanceFromReady.value = inst;
            onReady(inst);
            const passthrough = options.listeners?.onReady;
            if (passthrough) passthrough(inst);
          },
          ref: (el) => {
            childRef.value = (el as ComponentPublicInstance | null) ?? null;
          },
        });
    },
  });

  const app = createApp(Wrapper);
  app.mount(host);
  await flush();

  for (let i = 0; i < 20; i += 1) {
    if (instanceFromReady.value) break;
    await flush();
  }

  const instance = instanceFromReady.value;
  if (!instance) {
    throw new Error('mountVueSpreadsheet: instance never became available');
  }

  // Bridge: read the unwrapped `instance` off the exposed proxy each time
  // `.value` is accessed. This matches the public `SpreadsheetExposed`
  // signature (`Ref<SpreadsheetInstance | null>`) without requiring callers
  // to learn Vue's expose-unwrapping rules.
  const exposed: SpreadsheetExposed = {
    get instance() {
      const child = childRef.value as unknown as { instance: SpreadsheetInstance | null } | null;
      const cur = child?.instance ?? null;
      // Synthesize a Ref-like view so callers can use `.value`.
      return { value: cur } as unknown as SpreadsheetExposed['instance'];
    },
  };

  return {
    host,
    app,
    instance,
    exposed,
    async setProp(key, value) {
      propsState.value = { ...propsState.value, [key]: value };
      await flush();
    },
    async dispose() {
      app.unmount();
      await flush();
      host.remove();
      uninstallVueDomStubs();
    },
  };
}
