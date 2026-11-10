// SpreadsheetToolbar is a thin adapter over `mountToolbar`: it mounts the ribbon against its host,
// keeps the active tab in step in both directions, and forwards hooks at click time.

import type { RibbonTab, Toolbar } from '@libraz/formulon-cell/toolbar';
import { RIBBON_TABS } from '@libraz/formulon-cell/toolbar';
import { act, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { SpreadsheetToolbar, type SpreadsheetToolbarProps } from '../src/SpreadsheetToolbar';
import {
  installReactDomStubs,
  type MountedReactSpreadsheet,
  mountReactSpreadsheet,
  uninstallReactDomStubs,
} from './test-utils/mount';

const mountToolbarSpy = vi.hoisted(() => vi.fn());

vi.mock('@libraz/formulon-cell/toolbar', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@libraz/formulon-cell/toolbar')>();
  mountToolbarSpy.mockImplementation(actual.mountToolbar);
  return { ...actual, mountToolbar: mountToolbarSpy };
});

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
vi.setConfig({ testTimeout: 20_000 });

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

interface Harness {
  host: HTMLElement;
  root: Root;
  render: (node: ReactNode) => Promise<void>;
  unmount: () => Promise<void>;
}

const mountNode = async (node: ReactNode): Promise<Harness> => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  const render = async (next: ReactNode): Promise<void> => {
    await act(async () => {
      root.render(next);
      await flush();
    });
  };
  await render(node);
  return {
    host,
    root,
    render,
    unmount: async () => {
      await act(async () => {
        root.unmount();
        await flush();
      });
      host.remove();
    },
  };
};

describe('<SpreadsheetToolbar>', () => {
  let mounted: MountedReactSpreadsheet;
  let harness: Harness | null = null;

  beforeEach(async () => {
    installReactDomStubs();
    mounted = await mountReactSpreadsheet({ locale: 'en' });
    mountToolbarSpy.mockClear();
  });

  afterEach(async () => {
    if (harness) await harness.unmount();
    harness = null;
    await mounted.dispose();
    document.body.querySelectorAll('.fc-tb__dlg').forEach((el) => {
      el.remove();
    });
    uninstallReactDomStubs();
  });

  const toolbar = (props: Partial<SpreadsheetToolbarProps> = {}): ReactNode => (
    <SpreadsheetToolbar spreadsheet={mounted.instance} {...props} />
  );
  const click = async (selector: string): Promise<void> => {
    await act(async () => {
      harness?.host.querySelector<HTMLButtonElement>(selector)?.click();
      await flush();
    });
  };
  const visiblePanel = (tab: string): Element | null =>
    harness?.host.querySelector(`[data-ribbon-panel="${tab}"]:not([hidden])`) ?? null;

  it('mounts the ribbon into its host, hands out the toolbar and releases it on unmount', async () => {
    const onReady = vi.fn();
    harness = await mountNode(toolbar({ onReady }));

    expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    expect(visiblePanel('home')).toBeTruthy();
    const handed = onReady.mock.calls[0]?.[0] as Toolbar;
    expect(handed.disposed).toBe(false);
    expect(handed.activeTab).toBe('home');

    await harness.unmount();
    harness = null;
    expect(onReady).toHaveBeenLastCalledWith(null);
    expect(handed.disposed).toBe(true);
  });

  it('mounts nothing while the spreadsheet is null', async () => {
    harness = await mountNode(toolbar({ spreadsheet: null }));

    expect(mountToolbarSpy).not.toHaveBeenCalled();
    expect(harness.host.querySelector('[data-ribbon-tab]')).toBeNull();
  });

  it('reports a failed toolbar mount to onError', async () => {
    const err = new Error('toolbar mount failed');
    const onError = vi.fn();
    const onReady = vi.fn();
    mountToolbarSpy.mockRejectedValueOnce(err);
    harness = await mountNode(toolbar({ onError, onReady }));

    expect(onError).toHaveBeenCalledWith(err);
    expect(onReady).not.toHaveBeenCalled();
    expect(harness.host.querySelector('[data-ribbon-tab]')).toBeNull();
  });

  it('forwards tab-button clicks through onTabChange', async () => {
    const onTabChange = vi.fn();
    harness = await mountNode(toolbar({ onTabChange }));

    await click('[data-ribbon-tab="insert"]');

    expect(onTabChange).toHaveBeenCalledTimes(1);
    expect(onTabChange).toHaveBeenCalledWith('insert');
  });

  it('starts on defaultActiveTab and keeps the uncontrolled tab on user clicks', async () => {
    const onTabChange = vi.fn();
    harness = await mountNode(toolbar({ defaultActiveTab: 'data', onTabChange }));
    expect(visiblePanel('data')).toBeTruthy();

    await click('[data-ribbon-tab="insert"]');

    expect(visiblePanel('insert')).toBeTruthy();
    expect(onTabChange).toHaveBeenCalledWith('insert');
  });

  it('follows a controlled activeTab without remounting and does not echo it back', async () => {
    const onTabChange = vi.fn();
    harness = await mountNode(toolbar({ activeTab: 'home', onTabChange }));
    expect(visiblePanel('home')).toBeTruthy();
    expect(mountToolbarSpy).toHaveBeenCalledTimes(1);

    await harness.render(toolbar({ activeTab: 'data', onTabChange }));

    expect(visiblePanel('data')).toBeTruthy();
    expect(visiblePanel('home')).toBeFalsy();
    expect(mountToolbarSpy).toHaveBeenCalledTimes(1);
    expect(onTabChange).not.toHaveBeenCalled();
  });

  it('routes ribbon commands to the hooks current at click time', async () => {
    const spelling = vi.fn();
    const translate = vi.fn();
    const accessibility = vi.fn();
    const setInkMode = vi.fn();
    harness = await mountNode(
      toolbar({
        ribbonTabs: RIBBON_TABS,
        defaultActiveTab: 'review',
        hooks: { review: { spelling, translate, accessibility }, drawing: { setInkMode } },
      }),
    );

    await click('[data-ribbon-command="spellingReview"]');
    await click('[data-ribbon-command="translateReview"]');
    await click('[data-ribbon-command="accessibility"]');
    await click('[data-ribbon-command="drawPen"]');
    await click('[data-ribbon-command="drawErase"]');

    expect(spelling).toHaveBeenCalledTimes(1);
    expect(translate).toHaveBeenCalledTimes(1);
    expect(accessibility).toHaveBeenCalledTimes(1);
    expect(setInkMode.mock.calls).toEqual([['pen'], ['erase']]);
  });

  it('remounts only when a hook appears or disappears, not when its identity changes', async () => {
    const first = vi.fn();
    const second = vi.fn();
    const third = vi.fn();
    harness = await mountNode(toolbar({ defaultActiveTab: 'review' }));
    expect(mountToolbarSpy).toHaveBeenCalledTimes(1);

    await harness.render(
      toolbar({ defaultActiveTab: 'review', hooks: { review: { spelling: first } } }),
    );
    expect(mountToolbarSpy).toHaveBeenCalledTimes(2);

    await harness.render(
      toolbar({ defaultActiveTab: 'review', hooks: { review: { spelling: second } } }),
    );
    expect(mountToolbarSpy).toHaveBeenCalledTimes(2);
    await click('[data-ribbon-command="spellingReview"]');
    expect(first).not.toHaveBeenCalled();
    expect(second).toHaveBeenCalledTimes(1);

    await harness.render(
      toolbar({
        defaultActiveTab: 'review',
        hooks: { review: { spelling: second, translate: third } },
      }),
    );
    expect(mountToolbarSpy).toHaveBeenCalledTimes(3);
  });

  it('compares ribbonTabs structurally: an equal inline array keeps the ribbon, another set remounts', async () => {
    const tabs = (): RibbonTab[] => ['home', 'insert'];
    harness = await mountNode(toolbar({ ribbonTabs: tabs() }));
    expect(mountToolbarSpy).toHaveBeenCalledTimes(1);
    expect(harness.host.querySelectorAll('[data-ribbon-tab]')).toHaveLength(2);

    await harness.render(toolbar({ ribbonTabs: tabs() }));
    expect(mountToolbarSpy).toHaveBeenCalledTimes(1);

    await harness.render(toolbar({ ribbonTabs: ['home', 'insert', 'data'] }));
    expect(mountToolbarSpy).toHaveBeenCalledTimes(2);
    expect(harness.host.querySelectorAll('[data-ribbon-tab]')).toHaveLength(3);
  });

  it('remounts for another spreadsheet', async () => {
    harness = await mountNode(toolbar());
    const other = await mountReactSpreadsheet({ locale: 'en' });
    try {
      await harness.render(<SpreadsheetToolbar spreadsheet={other.instance} />);
      expect(mountToolbarSpy).toHaveBeenCalledTimes(2);
      expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
      await harness.unmount();
      harness = null;
    } finally {
      await other.dispose();
    }
  });
});
