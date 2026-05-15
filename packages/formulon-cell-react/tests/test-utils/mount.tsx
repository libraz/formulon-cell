import { type SpreadsheetInstance, WorkbookHandle } from '@libraz/formulon-cell';
import { act, type ReactElement, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { vi } from 'vitest';
import { Spreadsheet, type SpreadsheetProps, type SpreadsheetRef } from '../../src/Spreadsheet';

// React 18+ requires this global so its act() helper doesn't warn — set once
// at module load so any importer (Spreadsheet test, hooks test, …) gets it.
(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

/** Result of {@link mountReactSpreadsheet}. The caller should always call
 *  `dispose()` from an `afterEach` hook so the React root is unmounted, the
 *  host is detached from `document.body`, and the canvas / ResizeObserver
 *  stubs are restored.
 */
export interface MountedReactSpreadsheet {
  host: HTMLDivElement;
  root: Root;
  /** Live core instance after mount completes. Throws if mount failed. */
  instance: SpreadsheetInstance;
  /** The React ref forwarded to `<Spreadsheet>` so callers can re-use it. */
  refValue: SpreadsheetRef;
  /** Re-render with a new set of props. */
  rerender: (next: SpreadsheetProps, children?: ReactNode) => Promise<void>;
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

/** Install canvas + ResizeObserver stubs that let `Spreadsheet.mount` complete
 *  in happy-dom. Mirrors `tests/test-utils/dom.ts` in the core package — we
 *  inline rather than import across packages to keep this helper self-contained
 *  for the wrapper test runs.
 */
export const installReactDomStubs = (): void => {
  vi.stubGlobal('ResizeObserver', StubResizeObserver);
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(makeCanvasContext());
};

export const uninstallReactDomStubs = (): void => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
};

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

/** Mount the wrapper `<Spreadsheet>` against the real core (stub engine) and
 *  return a handle whose `instance` is guaranteed non-null. */
export async function mountReactSpreadsheet(
  props: SpreadsheetProps = {},
  children?: ReactNode,
): Promise<MountedReactSpreadsheet> {
  installReactDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  // Pre-create a workbook so the test never accidentally races a default-
  // workbook load that requires the real WASM engine.
  const workbook = props.workbook ?? (await WorkbookHandle.createDefault({ preferStub: true }));

  const refHolder: { current: SpreadsheetRef | null } = { current: null };
  const setRef = (next: SpreadsheetRef | null): void => {
    refHolder.current = next;
  };

  const renderTree = (next: SpreadsheetProps, kids?: ReactNode): ReactElement => (
    <Spreadsheet ref={setRef} workbook={workbook} {...next}>
      {kids}
    </Spreadsheet>
  );

  const onReady = vi.fn();
  await act(async () => {
    root.render(
      renderTree(
        {
          ...props,
          onReady: (i) => {
            onReady(i);
            props.onReady?.(i);
          },
        },
        children,
      ),
    );
    await flush();
  });

  // Wait for mount to complete (onReady fires after `Spreadsheet.mount` resolves).
  for (let i = 0; i < 20; i += 1) {
    if (onReady.mock.calls.length > 0) break;
    await act(async () => {
      await flush();
    });
  }

  const instance = refHolder.current?.instance;
  if (!instance) {
    throw new Error('mountReactSpreadsheet: instance never became available');
  }
  if (!refHolder.current) {
    throw new Error('mountReactSpreadsheet: ref never populated');
  }

  return {
    host,
    root,
    instance,
    refValue: refHolder.current,
    async rerender(next, kids) {
      await act(async () => {
        root.render(renderTree(next, kids));
        await flush();
      });
    },
    async dispose() {
      await act(async () => {
        root.unmount();
        await flush();
      });
      host.remove();
      uninstallReactDomStubs();
    },
  };
}
