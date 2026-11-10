import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { act, type ReactNode } from 'react';
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

  // biome-ignore lint/correctness/noComponentHookFactories: a probe built per render root hosts the hook under test and mounts once.
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

type Props = { instance: SpreadsheetInstance | null };

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
    const harness = renderHook((p: Props) => useSelection(p.instance), { instance: inst });
    await harness.rerender();

    expect(harness.value()?.active).toMatchObject({ row: 0, col: 0 });
    const startRenders = harness.renderCount();

    await act(async () => {
      inst.selection.set('E8');
      await flush();
    });

    expect(harness.value()?.active).toMatchObject({ row: 7, col: 4 });
    expect(harness.renderCount()).toBeGreaterThan(startRenders);
    await harness.unmount();
  });

  it('is null while the spreadsheet is null and stops listening once it is dropped', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook((p: Props) => useSelection(p.instance), {
      instance: inst as SpreadsheetInstance | null,
    });
    await harness.rerender();
    expect(harness.value()).not.toBeNull();

    await harness.setProps({ instance: null });
    expect(harness.value()).toBeNull();

    const before = harness.renderCount();
    await act(async () => {
      inst.selection.set('J10');
      await flush();
    });
    expect(harness.renderCount()).toBe(before);
    await harness.unmount();
  });

  it('is null once the spreadsheet is disposed', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook((p: Props) => useSelection(p.instance), { instance: inst });
    await harness.rerender();

    await mounted.dispose();
    mounted = null;
    await harness.rerender();

    expect(harness.value()).toBeNull();
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

  const a1 = (s: SpreadsheetInstance): number | null => {
    const v = s.workbook.sheet(0).getTypedValue(0, 0) as { number?: number };
    return v.number ?? null;
  };

  it('runs the selector against live state and re-runs on cell changes', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook((p: Props) => useSpreadsheet(p.instance, a1), { instance: inst });
    await harness.rerender();

    expect(harness.value()).toBeNull();

    await act(async () => {
      inst.workbook.sheet(0).setNumber(0, 0, 5);
      await flush();
    });

    expect(harness.value()).toBe(5);
    await harness.unmount();
  });

  it('re-runs the selector on selection and locale changes', async () => {
    mounted = await mountReactSpreadsheet({ locale: 'en' });
    const inst = mounted.instance;
    const harness = renderHook(
      (p: Props) =>
        useSpreadsheet(p.instance, (s) => `${s.selection.get().active.row}:${s.i18n.locale}`),
      { instance: inst },
    );
    await harness.rerender();
    expect(harness.value()).toBe('0:en');

    await act(async () => {
      inst.selection.set('A3');
      await flush();
    });
    expect(harness.value()).toBe('2:en');

    await act(async () => {
      inst.update({ locale: 'ja' });
      await flush();
    });
    expect(harness.value()).toBe('2:ja');
    await harness.unmount();
  });

  it('keeps fresh object selector results from causing a render loop', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook(
      (p: Props) => useSpreadsheet(p.instance, (s) => ({ row: s.selection.get().active.row })),
      { instance: inst },
    );
    await harness.rerender();

    expect(harness.value()).toEqual({ row: 0 });
    expect(harness.renderCount()).toBeLessThan(5);
    await harness.unmount();
  });

  it('is null while the spreadsheet is null and stops listening on unmount', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook((p: Props) => useSpreadsheet(p.instance, a1), {
      instance: null as SpreadsheetInstance | null,
    });
    await harness.rerender();
    expect(harness.value()).toBeNull();

    await harness.setProps({ instance: inst });
    inst.workbook.sheet(0).setNumber(0, 0, 3);
    await act(async () => {
      await flush();
    });
    expect(harness.value()).toBe(3);

    const before = harness.renderCount();
    await harness.unmount();
    inst.workbook.sheet(0).setNumber(0, 0, 4);
    await flush();
    expect(harness.renderCount()).toBe(before);
  });
});

describe('useI18n', () => {
  let mounted: MountedReactSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('mirrors the live locale and re-renders when the locale is updated', async () => {
    mounted = await mountReactSpreadsheet({ locale: 'ja' });
    const inst = mounted.instance;
    const harness = renderHook((p: Props) => useI18n(p.instance), { instance: inst });
    await harness.rerender();

    expect(harness.value()?.locale).toBe('ja');
    expect(harness.value()?.strings).toBe(inst.i18n.strings);

    await act(async () => {
      inst.update({ locale: 'en' });
      await flush();
    });

    expect(harness.value()?.locale).toBe('en');
    expect(harness.value()?.strings).toBe(inst.i18n.strings);
    await harness.unmount();
  });

  it('is null while the spreadsheet is null and does not re-render after unmount', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const harness = renderHook((p: Props) => useI18n(p.instance), {
      instance: null as SpreadsheetInstance | null,
    });
    await harness.rerender();
    expect(harness.value()).toBeNull();

    await harness.setProps({ instance: inst });
    expect(harness.value()?.locale).toBe(inst.i18n.locale);

    const beforeUnmount = harness.renderCount();
    await harness.unmount();

    inst.update({ locale: 'ja' });
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

  interface ProbeProps {
    instance: SpreadsheetInstance | null;
    handler: (e: unknown) => void;
  }
  // biome-ignore lint/correctness/noComponentHookFactories: a probe shared by this block hosts the hook under test.
  const Probe = ({ instance, handler }: ProbeProps): ReactNode => {
    useSpreadsheetEvent(instance, 'cell', handler);
    return null;
  };

  const render = async (props: ProbeProps): Promise<void> => {
    if (!outerRoot) {
      const host = document.createElement('div');
      document.body.appendChild(host);
      outerRoot = createRoot(host);
    }
    await act(async () => {
      outerRoot?.render(<Probe {...props} />);
      await flush();
    });
  };

  it('subscribes to the named event and invokes the latest handler without resubscribing', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const on = vi.spyOn(inst, 'on');
    const handlerA = vi.fn();
    const handlerB = vi.fn();

    await render({ instance: inst, handler: handlerA });
    inst.workbook.sheet(0).setNumber(0, 0, 1);
    expect(handlerA).toHaveBeenCalledTimes(1);
    expect(handlerA.mock.calls[0]?.[0]).toMatchObject({ sheet: 0, at: { row: 0, col: 0 } });

    await render({ instance: inst, handler: handlerB });
    inst.workbook.sheet(0).setNumber(1, 0, 2);
    expect(handlerB).toHaveBeenCalledTimes(1);
    expect(handlerA).toHaveBeenCalledTimes(1);
    expect(on).toHaveBeenCalledTimes(1);
  });

  it('cleans up the subscription on unmount so the handler stops firing', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const handler = vi.fn();

    await render({ instance: inst, handler });
    inst.workbook.sheet(0).setNumber(0, 0, 1);
    expect(handler).toHaveBeenCalledTimes(1);

    await act(async () => {
      outerRoot?.unmount();
      await flush();
    });
    outerRoot = null;

    handler.mockClear();
    inst.workbook.sheet(0).setNumber(0, 1, 2);
    await flush();
    expect(handler).not.toHaveBeenCalled();
  });

  it('does nothing while the spreadsheet is null', async () => {
    const handler = vi.fn();
    await render({ instance: null, handler });
    expect(handler).not.toHaveBeenCalled();
  });
});
