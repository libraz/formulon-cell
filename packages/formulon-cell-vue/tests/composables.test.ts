import { createSpreadsheetStore, mutators, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  type App,
  type ComponentPublicInstance,
  createApp,
  defineComponent,
  nextTick,
  type Ref,
  shallowRef,
} from 'vue';
import { useI18n, useSelection, useSpreadsheet, useSpreadsheetEvent } from '../src/composables';
import {
  installVueDomStubs,
  type MountedVueSpreadsheet,
  mountVueSpreadsheet,
  uninstallVueDomStubs,
} from './test-utils/mount';

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
  await nextTick();
};

interface ComposableHarness<T> {
  /** Latest composable return value. */
  readonly value: () => T;
  /** Update the wrapped instance ref the composable depends on. */
  setInstance(next: SpreadsheetInstance | null): Promise<void>;
  unmount(): Promise<void>;
}

/** Mount a tiny Vue app whose `setup()` calls the supplied composable and
 *  exposes its return value via a captured ref. */
function renderComposable<T>(
  factory: (instance: Ref<SpreadsheetInstance | null>) => T,
  initial: SpreadsheetInstance | null,
): ComposableHarness<T> & { app: App } {
  installVueDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);

  const instanceRef = shallowRef<SpreadsheetInstance | null>(initial);
  let captured: T | undefined;
  const Probe = defineComponent({
    setup() {
      captured = factory(instanceRef);
      return () => null;
    },
  });

  const app = createApp(Probe);
  app.mount(host);

  return {
    app,
    value: () => {
      if (captured === undefined) throw new Error('composable never produced a value');
      return captured;
    },
    async setInstance(next) {
      instanceRef.value = next;
      await flush();
    },
    async unmount() {
      app.unmount();
      await flush();
      host.remove();
      uninstallVueDomStubs();
    },
  };
}

describe('useSelection', () => {
  let mounted: MountedVueSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('mirrors the live selection and updates when the active cell moves', async () => {
    mounted = await mountVueSpreadsheet();
    const inst = mounted.instance;
    const harness = renderComposable((r) => useSelection(r), inst);
    await flush();

    expect(harness.value().value.active).toEqual({ sheet: 0, row: 0, col: 0 });

    mutators.setActive(inst.store, { sheet: 0, row: 5, col: 2 });
    await flush();

    expect(harness.value().value.active).toEqual({ sheet: 0, row: 5, col: 2 });

    await harness.unmount();
  });

  it('cleans up its store subscription on unmount', async () => {
    mounted = await mountVueSpreadsheet();
    const inst = mounted.instance;
    const harness = renderComposable((r) => useSelection(r), inst);
    await flush();

    const refValue = harness.value();
    const before = refValue.value.active;
    expect(before).toEqual({ sheet: 0, row: 0, col: 0 });

    await harness.unmount();

    // Mutating after unmount must not throw and must not crash subscribers.
    mutators.setActive(inst.store, { sheet: 0, row: 6, col: 6 });
    await flush();
    // The captured ref's last value reflects pre-unmount state — proves the
    // subscription was severed (otherwise it'd track the mutation).
    expect(refValue.value.active).toEqual(before);
  });

  it('resets on null, follows a new instance, and ignores the old instance', async () => {
    mounted = await mountVueSpreadsheet();
    const first = mounted.instance;
    const otherStore = createSpreadsheetStore();
    const other = { store: otherStore } as unknown as SpreadsheetInstance;
    const harness = renderComposable((r) => useSelection(r), first);
    await flush();

    mutators.setActive(first.store, { sheet: 0, row: 2, col: 3 });
    await flush();
    await harness.setInstance(null);
    expect(harness.value().value.active).toEqual({ sheet: 0, row: 0, col: 0 });

    await harness.setInstance(other);
    mutators.setActive(otherStore, { sheet: 0, row: 4, col: 5 });
    await flush();
    expect(harness.value().value.active).toEqual({ sheet: 0, row: 4, col: 5 });

    mutators.setActive(first.store, { sheet: 0, row: 8, col: 9 });
    await flush();
    expect(harness.value().value.active).toEqual({ sheet: 0, row: 4, col: 5 });

    await harness.unmount();
  });

  it('does not invoke a detached store unsubscribe again on scope dispose', async () => {
    const store = createSpreadsheetStore();
    const off = vi.fn();
    const fake = {
      store: {
        getState: () => store.getState(),
        subscribe: () => off,
      },
    } as unknown as SpreadsheetInstance;
    const harness = renderComposable((r) => useSelection(r), fake);
    await flush();

    await harness.setInstance(null);
    await harness.unmount();

    expect(off).toHaveBeenCalledTimes(1);
  });
});

describe('useSpreadsheet', () => {
  let mounted: MountedVueSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('runs the selector against the live store and re-runs on changes', async () => {
    mounted = await mountVueSpreadsheet();
    const inst = mounted.instance;

    const harness = renderComposable((r) => useSpreadsheet(r, (s) => s.data.cells.size, -1), inst);
    await flush();

    expect(harness.value().value).toBe(0);

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 7);
    inst.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 8);
    await flush();

    expect(harness.value().value).toBe(2);
    await harness.unmount();
  });

  it('returns the supplied fallback when instance is null and stops listening on unmount', async () => {
    mounted = await mountVueSpreadsheet();
    const inst = mounted.instance;

    const harness = renderComposable(
      (r) => useSpreadsheet<number | 'fallback'>(r, (s) => s.data.cells.size, 'fallback'),
      null,
    );
    await flush();

    expect(harness.value().value).toBe('fallback');

    await harness.setInstance(inst);
    expect(harness.value().value).toBe(0);

    await harness.unmount();

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    await flush();
    // The stored ref doesn't update post-unmount.
    expect(harness.value().value).toBe(0);
  });
});

describe('useI18n', () => {
  let mounted: MountedVueSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('mirrors the live locale and updates when setLocale fires', async () => {
    mounted = await mountVueSpreadsheet({ props: { locale: 'ja' } });
    const inst = mounted.instance;

    const harness = renderComposable((r) => useI18n(r), inst);
    await flush();

    expect(harness.value().locale.value).toBe('ja');
    expect(harness.value().strings.value).toBeTruthy();

    inst.i18n.setLocale('en');
    await flush();

    expect(harness.value().locale.value).toBe('en');
    await harness.unmount();
  });

  it('cleans up its i18n subscription on unmount', async () => {
    mounted = await mountVueSpreadsheet({ props: { locale: 'ja' } });
    const inst = mounted.instance;

    const harness = renderComposable((r) => useI18n(r), inst);
    await flush();

    const localeRef = harness.value().locale;
    expect(localeRef.value).toBe('ja');

    await harness.unmount();

    // Mutating the locale after unmount must not propagate to the captured
    // ref (proves the subscription was severed).
    inst.i18n.setLocale('en');
    await flush();
    expect(localeRef.value).toBe('ja');
  });

  it('resets on null, follows a new instance, and ignores the old instance', async () => {
    mounted = await mountVueSpreadsheet({ props: { locale: 'ja' } });
    const first = mounted.instance;
    let otherLocale = 'en';
    const otherStrings = first.i18n.strings;
    const listeners = new Set<(strings: typeof otherStrings) => void>();
    const other = {
      i18n: {
        get locale() {
          return otherLocale;
        },
        get strings() {
          return otherStrings;
        },
        subscribe(listener: (strings: typeof otherStrings) => void) {
          listeners.add(listener);
          return () => listeners.delete(listener);
        },
      },
    } as unknown as SpreadsheetInstance;
    const harness = renderComposable((r) => useI18n(r), first);
    await flush();

    first.i18n.setLocale('en');
    await flush();
    await harness.setInstance(null);
    expect(harness.value().locale.value).toBe('ja');
    expect(harness.value().strings.value).toEqual({});

    await harness.setInstance(other);
    expect(harness.value().locale.value).toBe('en');
    otherLocale = 'fr';
    for (const listener of listeners) listener(otherStrings);
    await flush();
    expect(harness.value().locale.value).toBe('fr');

    first.i18n.setLocale('ja');
    await flush();
    expect(harness.value().locale.value).toBe('fr');

    await harness.unmount();
  });

  it('does not invoke a detached i18n unsubscribe again on scope dispose', async () => {
    const strings = {} as SpreadsheetInstance['i18n']['strings'];
    const off = vi.fn();
    const fake = {
      i18n: {
        locale: 'ja',
        strings,
        subscribe: () => off,
      },
    } as unknown as SpreadsheetInstance;
    const harness = renderComposable((r) => useI18n(r), fake);
    await flush();

    await harness.setInstance(null);
    await harness.unmount();

    expect(off).toHaveBeenCalledTimes(1);
  });
});

describe('useSpreadsheetEvent', () => {
  let mounted: MountedVueSpreadsheet | null = null;
  let helperApp: { app: App; host: HTMLElement } | null = null;

  afterEach(async () => {
    if (helperApp) {
      helperApp.app.unmount();
      helperApp.host.remove();
      uninstallVueDomStubs();
      helperApp = null;
    }
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('subscribes to the named event and forwards payloads to the handler', async () => {
    mounted = await mountVueSpreadsheet();
    const inst = mounted.instance;

    const handler = vi.fn();
    installVueDomStubs();
    const host = document.createElement('div');
    document.body.appendChild(host);

    const instanceRef = shallowRef<SpreadsheetInstance | null>(inst);
    const Probe = defineComponent({
      setup() {
        useSpreadsheetEvent(instanceRef, 'cellChange', handler);
        return () => null;
      },
    });
    const app = createApp(Probe);
    app.mount(host);
    helperApp = { app, host };
    await flush();

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 9);
    await flush();

    expect(handler).toHaveBeenCalledTimes(1);
    const event = handler.mock.calls[0]?.[0];
    expect((event as { value: { kind: string; value: number } } | undefined)?.value).toEqual({
      kind: 'number',
      value: 9,
    });
  });

  it('cleans up on unmount and stops invoking the handler afterwards', async () => {
    mounted = await mountVueSpreadsheet();
    const inst = mounted.instance;

    const handler = vi.fn();
    installVueDomStubs();
    const host = document.createElement('div');
    document.body.appendChild(host);

    const instanceRef = shallowRef<SpreadsheetInstance | null>(inst);
    const Probe = defineComponent({
      setup() {
        useSpreadsheetEvent(instanceRef, 'cellChange', handler);
        return () => null;
      },
    });
    const app = createApp(Probe);
    const probeMount = app.mount(host) as ComponentPublicInstance;
    expect(probeMount).toBeDefined();
    await flush();

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    await flush();
    expect(handler).toHaveBeenCalledTimes(1);

    app.unmount();
    host.remove();
    uninstallVueDomStubs();

    handler.mockClear();
    inst.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    await flush();
    expect(handler).not.toHaveBeenCalled();
  });
});
