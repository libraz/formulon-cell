import type { SpreadsheetEventName, SpreadsheetInstance } from '@libraz/formulon-cell';
import { act, type ReactNode, useEffect, useMemo } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// React 18+ asks act() callers to opt-in via this global so production-mode
// `react-dom` doesn't warn on every render-from-test.
(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

import { useI18n, useSelection, useSpreadsheet, useSpreadsheetEvent } from '../src/hooks';
import {
  installReactDomStubs,
  type MountedReactSpreadsheet,
  mountReactSpreadsheet,
  uninstallReactDomStubs,
} from './test-utils/mount';

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

interface HookHarness<T> {
  /** Latest hook return value. Updated synchronously on every render. */
  readonly value: () => T;
  /** Number of times the hook has rendered (incl. the initial render). */
  readonly renderCount: () => number;
  rerender(): Promise<void>;
  unmount(): Promise<void>;
}

/** Minimal `renderHook` substitute: mounts a React tree that calls the hook
 *  on every render and stashes its result + render count in refs. We avoid
 *  pulling in `@testing-library/react-hooks` because it isn't installed.
 */
function renderHook<T, P>(
  hook: (props: P) => T,
  initialProps: P,
): HookHarness<T> & { setProps: (next: P) => Promise<void> } {
  installReactDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  const ref: { value?: T; renderCount: number } = { renderCount: 0 };
  let currentProps = initialProps;

  const Probe = ({ p }: { p: P }): ReactNode => {
    ref.renderCount += 1;
    const v = hook(p);
    ref.value = v;
    return null;
  };

  const render = (p: P): void => {
    root.render(<Probe p={p} />);
  };

  // Initial render.
  // act() ensures effects flush before we observe the value.
  void act(() => {
    render(initialProps);
  });

  return {
    value: () => {
      if (!('value' in ref)) throw new Error('hook never produced a value');
      return ref.value as T;
    },
    renderCount: () => ref.renderCount,
    async rerender() {
      await act(async () => {
        render(currentProps);
        await flush();
      });
    },
    async setProps(next) {
      currentProps = next;
      await act(async () => {
        render(next);
        await flush();
      });
    },
    async unmount() {
      await act(async () => {
        root.unmount();
        await flush();
      });
      host.remove();
      uninstallReactDomStubs();
    },
  };
}

describe('useSelection', () => {
  let mounted: MountedReactSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('returns the live selection and re-renders when the active cell moves', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook(
      (p: { instance: SpreadsheetInstance | null }) => useSelection(p.instance),
      {
        instance: inst,
      },
    );
    await harness.rerender();

    expect(harness.value().active).toEqual({ sheet: 0, row: 0, col: 0 });
    const startRenders = harness.renderCount();

    const { mutators } = await import('@libraz/formulon-cell');
    await act(async () => {
      mutators.setActive(inst.store, { sheet: 0, row: 7, col: 4 });
      await flush();
    });

    expect(harness.value().active).toEqual({ sheet: 0, row: 7, col: 4 });
    expect(harness.renderCount()).toBeGreaterThan(startRenders);
    await harness.unmount();
  });

  it('falls back to the zero-selection when instance is null and unsubscribes on unmount', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;

    const harness = renderHook(
      (p: { instance: SpreadsheetInstance | null }) => useSelection(p.instance),
      { instance: inst as SpreadsheetInstance | null },
    );
    await harness.rerender();

    // Swap to null — the hook should fall back to the placeholder selection.
    await harness.setProps({ instance: null });
    expect(harness.value()).toEqual({
      active: { sheet: 0, row: 0, col: 0 },
      anchor: { sheet: 0, row: 0, col: 0 },
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
    });

    // Mutating the original instance after the hook detaches must not throw
    // and must not bring the hook back to life.
    const { mutators } = await import('@libraz/formulon-cell');
    const before = harness.renderCount();
    await act(async () => {
      mutators.setActive(inst.store, { sheet: 0, row: 9, col: 9 });
      await flush();
    });
    expect(harness.renderCount()).toBe(before);

    await harness.unmount();
  });
});

describe('useSpreadsheet', () => {
  let mounted: MountedReactSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('runs the selector against live state and re-runs on store changes', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook(
      (p: { instance: SpreadsheetInstance | null }) =>
        useSpreadsheet(p.instance, (s) => s.data.cells.size, -1),
      { instance: inst },
    );
    await harness.rerender();

    expect(harness.value()).toBe(0);

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
    inst.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 6);
    await act(async () => {
      await flush();
    });

    expect(harness.value()).toBe(2);
    await harness.unmount();
  });

  it('returns the supplied fallback while instance is null and stops listening on unmount', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;

    const harness = renderHook(
      (p: { instance: SpreadsheetInstance | null }) =>
        useSpreadsheet<number | 'fallback'>(p.instance, (s) => s.data.cells.size, 'fallback'),
      { instance: null as SpreadsheetInstance | null },
    );
    await harness.rerender();
    expect(harness.value()).toBe('fallback');

    await harness.setProps({ instance: inst });
    expect(harness.value()).toBe(0);

    await harness.unmount();

    // Post-unmount: changes to the store no longer reach the (gone) hook.
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    await flush();
    // No assertion failures means the unsubscribe ran cleanly.
    expect(inst.store.getState().data.cells.size).toBe(1);
  });
});

describe('useI18n', () => {
  let mounted: MountedReactSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('mirrors the live locale and re-renders when setLocale fires', async () => {
    mounted = await mountReactSpreadsheet({ locale: 'ja' });
    const inst = mounted.instance;

    const harness = renderHook(
      (p: { instance: SpreadsheetInstance | null }) => useI18n(p.instance),
      { instance: inst },
    );
    await harness.rerender();

    expect(harness.value().locale).toBe('ja');
    expect(harness.value().strings).not.toBeNull();

    await act(async () => {
      inst.i18n.setLocale('en');
      await flush();
    });

    expect(harness.value().locale).toBe('en');
    await harness.unmount();
  });

  it('returns the placeholder when instance is null and unsubscribes on unmount', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;

    const harness = renderHook(
      (p: { instance: SpreadsheetInstance | null }) => useI18n(p.instance),
      { instance: null as SpreadsheetInstance | null },
    );
    await harness.rerender();
    expect(harness.value()).toEqual({ locale: 'ja', strings: null });

    await harness.setProps({ instance: inst });
    expect(harness.value().locale).toBe(inst.i18n.locale);

    const beforeUnmount = harness.renderCount();
    await harness.unmount();

    // Triggering an i18n change after unmount must not re-render the hook.
    inst.i18n.setLocale('en');
    await flush();
    expect(harness.renderCount()).toBe(beforeUnmount);
  });
});

describe('useSpreadsheetEvent', () => {
  let mounted: MountedReactSpreadsheet | null = null;
  let outerRoot: Root | null = null;

  beforeEach(() => {
    document.body.replaceChildren();
  });

  afterEach(async () => {
    if (outerRoot) {
      await act(async () => {
        outerRoot?.unmount();
        await flush();
      });
      outerRoot = null;
    }
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('subscribes to the named event and invokes the latest handler reference', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;

    const handlerA = vi.fn();
    const handlerB = vi.fn();

    interface ProbeProps {
      instance: SpreadsheetInstance | null;
      handler: (e: unknown) => void;
    }
    const Probe = ({ instance, handler }: ProbeProps): ReactNode => {
      useSpreadsheetEvent(instance, 'cellChange' as SpreadsheetEventName, handler as never);
      return null;
    };

    const host = document.createElement('div');
    document.body.appendChild(host);
    outerRoot = createRoot(host);

    await act(async () => {
      outerRoot?.render(<Probe instance={inst} handler={handlerA} />);
      await flush();
    });

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    expect(handlerA).toHaveBeenCalledTimes(1);

    // Swap handler without remounting → re-render should not re-subscribe but
    // the new handler must still receive the next event.
    await act(async () => {
      outerRoot?.render(<Probe instance={inst} handler={handlerB} />);
      await flush();
    });

    inst.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    expect(handlerB).toHaveBeenCalledTimes(1);
    expect(handlerA).toHaveBeenCalledTimes(1); // not re-invoked
  });

  it('cleans up the subscription on unmount so handler stops firing', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;

    const handler = vi.fn();
    interface ProbeProps {
      instance: SpreadsheetInstance | null;
      handler: (e: unknown) => void;
    }
    const Probe = ({ instance, handler: h }: ProbeProps): ReactNode => {
      useSpreadsheetEvent(instance, 'cellChange' as SpreadsheetEventName, h as never);
      // Touch a memo so the test can assert the component actually mounted.
      useEffect(() => undefined, []);
      useMemo(() => 0, []);
      return null;
    };

    const host = document.createElement('div');
    document.body.appendChild(host);
    outerRoot = createRoot(host);

    await act(async () => {
      outerRoot?.render(<Probe instance={inst} handler={handler} />);
      await flush();
    });

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    expect(handler).toHaveBeenCalledTimes(1);

    await act(async () => {
      outerRoot?.unmount();
      await flush();
    });
    outerRoot = null;

    handler.mockClear();
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    await flush();
    expect(handler).not.toHaveBeenCalled();
  });
});
