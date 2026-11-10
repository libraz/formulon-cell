import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { RIBBON_TABS, type RibbonHooks, type RibbonTab } from '@libraz/formulon-cell/toolbar';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { type App, createApp, defineComponent, h, shallowReactive } from 'vue';
import { SpreadsheetToolbar } from '../src';
import { flush, mountStub, waitFor } from './test-utils/mount';

interface Harness {
  host: HTMLElement;
  props: Record<string, unknown>;
  ready: ReturnType<typeof vi.fn>;
  error: ReturnType<typeof vi.fn>;
  tab: ReturnType<typeof vi.fn>;
  unmount: () => void;
}

let instance: SpreadsheetInstance;
const apps: App[] = [];

beforeEach(async () => {
  instance = await mountStub();
});

afterEach(() => {
  for (const app of apps.splice(0)) app.unmount();
  instance.dispose();
  document.body.replaceChildren();
});

/** Mounts `<SpreadsheetToolbar>` with reactive props; `activeTab` is bound two-way like `v-model`. */
const mountToolbar = async (
  initial: Record<string, unknown> = {},
  { wait = true }: { wait?: boolean } = {},
): Promise<Harness> => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const props = shallowReactive<Record<string, unknown>>({ spreadsheet: instance, ...initial });
  const ready = vi.fn();
  const error = vi.fn();
  const tab = vi.fn();
  const app = createApp(
    defineComponent({
      render: () =>
        h(SpreadsheetToolbar, {
          ...props,
          onReady: ready,
          onError: error,
          'onUpdate:activeTab': (next: RibbonTab) => {
            tab(next);
            props.activeTab = next;
          },
        }),
    }),
  );
  app.mount(host);
  apps.push(app);
  if (wait && props.spreadsheet) await waitFor(() => ready.mock.calls.length > 0);
  return { host, props, ready, error, tab, unmount: () => app.unmount() };
};

const panel = (host: HTMLElement, name: string): Element | null =>
  host.querySelector(`[data-ribbon-panel="${name}"]:not([hidden])`);

describe('<SpreadsheetToolbar>', () => {
  it('mounts the ribbon into its host and emits ready with the toolbar', async () => {
    const view = await mountToolbar();
    expect(view.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    expect(panel(view.host, 'home')).toBeTruthy();
    expect(view.ready).toHaveBeenCalledTimes(1);
    expect(view.ready.mock.calls[0]?.[0]).toMatchObject({ activeTab: 'home' });
  });

  it('mounts nothing while the spreadsheet is null, then mounts when it arrives', async () => {
    const view = await mountToolbar({ spreadsheet: null });
    await flush();
    expect(view.host.querySelector('[data-ribbon-tab]')).toBeNull();
    expect(view.ready).not.toHaveBeenCalled();
    view.props.spreadsheet = instance;
    await waitFor(() => view.ready.mock.calls.length > 0);
    expect(view.host.querySelector('[data-ribbon-tab]')).not.toBeNull();
  });

  it('emits update:activeTab when a tab button is clicked', async () => {
    const view = await mountToolbar();
    view.host.querySelector<HTMLButtonElement>('[data-ribbon-tab="insert"]')?.click();
    await flush();
    expect(view.tab).toHaveBeenCalledWith('insert');
    expect(panel(view.host, 'insert')).toBeTruthy();
  });

  it('applies an external activeTab change without remounting', async () => {
    const view = await mountToolbar();
    view.props.activeTab = 'data';
    await flush();
    expect(panel(view.host, 'data')).toBeTruthy();
    expect(panel(view.host, 'home')).toBeFalsy();
    expect(view.ready).toHaveBeenCalledTimes(1);
  });

  it('does not echo an externally set tab back as update:activeTab', async () => {
    const view = await mountToolbar();
    view.props.activeTab = 'data';
    await flush();
    expect(view.tab).not.toHaveBeenCalled();
  });

  it('starts on the given activeTab', async () => {
    const view = await mountToolbar({ activeTab: 'review', ribbonTabs: RIBBON_TABS });
    expect(panel(view.host, 'review')).toBeTruthy();
  });

  it('emits error when the toolbar mount fails and mounts nothing', async () => {
    const disposed = await mountStub();
    disposed.dispose();
    const view = await mountToolbar({ spreadsheet: disposed }, { wait: false });
    await waitFor(() => view.error.mock.calls.length > 0);
    expect(view.ready).not.toHaveBeenCalled();
    expect(view.host.querySelector('[data-ribbon-tab]')).toBeNull();
  });

  it('emits ready(null) and clears the ribbon when unmounted', async () => {
    const view = await mountToolbar();
    const toolbar = view.ready.mock.calls[0]?.[0] as { disposed: boolean };
    view.unmount();
    expect(toolbar.disposed).toBe(true);
    expect(view.ready).toHaveBeenLastCalledWith(null);
  });

  describe('hooks', () => {
    const clickPen = async (host: HTMLElement): Promise<void> => {
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="drawPen"]')?.click();
      await flush();
    };
    const hooksOf = (setInkMode: () => void): RibbonHooks => ({ drawing: { setInkMode } });

    it('calls the hook given for a ribbon command', async () => {
      const setInkMode = vi.fn();
      const view = await mountToolbar({
        activeTab: 'review',
        ribbonTabs: RIBBON_TABS,
        hooks: hooksOf(setInkMode),
      });
      await clickPen(view.host);
      expect(setInkMode).toHaveBeenCalledWith('pen');
    });

    it('calls the latest function at click time and keeps the ribbon mounted', async () => {
      const first = vi.fn();
      const second = vi.fn();
      const view = await mountToolbar({
        activeTab: 'review',
        ribbonTabs: RIBBON_TABS,
        hooks: hooksOf(first),
      });
      view.props.hooks = hooksOf(second);
      await flush();
      expect(view.ready).toHaveBeenCalledTimes(1);
      await clickPen(view.host);
      expect(first).not.toHaveBeenCalled();
      expect(second).toHaveBeenCalledWith('pen');
    });

    it('remounts when a hook is added or removed', async () => {
      const view = await mountToolbar({ activeTab: 'review', ribbonTabs: RIBBON_TABS });
      const setInkMode = vi.fn();
      view.props.hooks = hooksOf(setInkMode);
      await waitFor(() => view.ready.mock.calls.length >= 3);
      expect(view.ready).toHaveBeenNthCalledWith(2, null);
      await clickPen(view.host);
      expect(setInkMode).toHaveBeenCalledWith('pen');
    });
  });

  describe('ribbonTabs', () => {
    it('does not remount for a structurally equal array', async () => {
      const view = await mountToolbar({ ribbonTabs: ['home', 'insert'] });
      view.props.ribbonTabs = ['home', 'insert'];
      await flush();
      expect(view.ready).toHaveBeenCalledTimes(1);
    });

    it('remounts with the new tab surface when the tabs differ', async () => {
      const view = await mountToolbar({ ribbonTabs: ['home', 'insert'] });
      expect(view.host.querySelector('[data-ribbon-tab="data"]')).toBeNull();
      view.props.ribbonTabs = ['home', 'insert', 'data'];
      await waitFor(() => view.host.querySelector('[data-ribbon-tab="data"]') !== null);
    });
  });

  it('remounts for another spreadsheet and disposes the previous toolbar', async () => {
    const view = await mountToolbar();
    const first = view.ready.mock.calls[0]?.[0] as { disposed: boolean };
    const other = await mountStub();
    view.props.spreadsheet = other;
    await waitFor(() => view.ready.mock.calls.filter((c) => c[0] !== null).length >= 2);
    expect(first.disposed).toBe(true);
    expect(view.host.querySelectorAll('[data-ribbon-tab="home"]').length).toBe(1);
    view.unmount();
    other.dispose();
  });
});
