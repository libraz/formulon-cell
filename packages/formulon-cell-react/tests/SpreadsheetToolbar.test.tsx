import {
  EMPTY_ACTIVE_STATE,
  mutators,
  projectActiveState,
  RIBBON_TAB_LABELS,
  type RibbonTab,
} from '@libraz/formulon-cell';
import { act, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { SpreadsheetToolbar } from '../src/SpreadsheetToolbar';
import {
  installReactDomStubs,
  type MountedReactSpreadsheet,
  mountReactSpreadsheet,
  uninstallReactDomStubs,
} from './test-utils/mount';

// React 18+ asks act() callers to opt-in via this global.
(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

interface ToolbarHarness {
  host: HTMLElement;
  root: Root;
  rerender(activeTab: RibbonTab, onTabChange?: (tab: RibbonTab) => void): Promise<void>;
  unmount(): Promise<void>;
}

async function renderToolbar(
  mounted: MountedReactSpreadsheet,
  initial: { activeTab: RibbonTab; onTabChange: (tab: RibbonTab) => void; locale?: string },
): Promise<ToolbarHarness> {
  installReactDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);

  const render = (activeTab: RibbonTab, onTabChange: (tab: RibbonTab) => void): ReactNode => (
    <SpreadsheetToolbar
      instance={mounted.instance}
      activeTab={activeTab}
      onTabChange={onTabChange}
      locale={initial.locale ?? 'en'}
    />
  );

  await act(async () => {
    root.render(render(initial.activeTab, initial.onTabChange));
    await flush();
  });

  let lastTab = initial.activeTab;
  let lastChange = initial.onTabChange;

  return {
    host,
    root,
    async rerender(activeTab, onTabChange) {
      lastTab = activeTab;
      if (onTabChange) lastChange = onTabChange;
      await act(async () => {
        root.render(render(lastTab, lastChange));
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

describe('React <SpreadsheetToolbar>', () => {
  let mounted: MountedReactSpreadsheet | null = null;
  let toolbar: ToolbarHarness | null = null;

  beforeEach(() => {
    document.body.replaceChildren();
  });

  afterEach(async () => {
    if (toolbar) {
      await toolbar.unmount();
      toolbar = null;
    }
    if (mounted) {
      await mounted.dispose();
      mounted = null;
    }
    document.body.replaceChildren();
  });

  it('renders a tab button for every non-file ribbon tab using RIBBON_TAB_LABELS', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const tabButtons = toolbar.host.querySelectorAll<HTMLButtonElement>('[role="tab"]');
    const labels = Array.from(tabButtons).map((b) => b.textContent?.trim());
    const expected = (Object.keys(RIBBON_TAB_LABELS) as RibbonTab[])
      .filter((id) => id !== 'file')
      .map((id) => RIBBON_TAB_LABELS[id].en);
    expect(labels).toEqual(expected);
  });

  it('invokes onTabChange when a tab button is clicked', async () => {
    mounted = await mountReactSpreadsheet();
    const onTabChange = vi.fn();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange });

    const tabs = toolbar.host.querySelectorAll<HTMLButtonElement>('[role="tab"]');
    const insertTab = Array.from(tabs).find((t) => t.textContent?.includes('Insert'));
    expect(insertTab).toBeDefined();

    await act(async () => {
      insertTab?.click();
      await flush();
    });

    expect(onTabChange).toHaveBeenCalledWith('insert');
  });

  it('marks the active tab with aria-selected=true and a CSS modifier', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const tabs = toolbar.host.querySelectorAll<HTMLButtonElement>('[role="tab"]');
    const dataTab = Array.from(tabs).find((t) => t.textContent?.includes('Data'));
    expect(dataTab?.getAttribute('aria-selected')).toBe('true');
    expect(dataTab?.className).toContain('demo__ribbon-tab--active');

    // A non-active tab should not carry the modifier.
    const homeTab = Array.from(tabs).find((t) => t.textContent?.includes('Home'));
    expect(homeTab?.getAttribute('aria-selected')).toBe('false');
    expect(homeTab?.className).not.toContain('demo__ribbon-tab--active');
  });

  it('reflects bold=true in the toolbar after the bold button is clicked', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    // Sanity: starts unbold.
    const before = projectActiveState(mounted.instance);
    expect(before.bold).toBe(EMPTY_ACTIVE_STATE.bold);
    expect(before.bold).toBe(false);

    // The bold button is the one whose aria-label starts with "Bold". Multiple
    // tabs may render bold-related controls; on Home the aria-label is
    // "Bold (⌘B)".
    const boldButton = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (b) => b.getAttribute('aria-label')?.startsWith('Bold'),
    );
    expect(boldButton).toBeDefined();

    await act(async () => {
      boldButton?.click();
      await flush();
    });

    const inst = mounted.instance;
    const after = projectActiveState(inst);
    expect(after.bold).toBe(true);

    // The button itself should now carry the active modifier — the toolbar
    // store-subscription has to fire before the next render reflects this.
    await act(async () => {
      mutators.setActive(inst.store, inst.store.getState().selection.active);
      await flush();
    });
    const refreshed = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (b) => b.getAttribute('aria-label')?.startsWith('Bold'),
    );
    expect(refreshed?.className).toContain('demo__rb--active');
  });

  it('unsubscribes from the store on unmount and ignores subsequent state changes silently', async () => {
    mounted = await mountReactSpreadsheet();
    const onTabChange = vi.fn();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange });

    // Capture console.error so we can assert nothing leaks after unmount.
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});

    await toolbar.unmount();
    toolbar = null;

    // Toggle a few store fields — the unmounted toolbar must not react.
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 9, col: 9 });
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 4 });
    await flush();

    expect(errSpy).not.toHaveBeenCalled();
    errSpy.mockRestore();
  });
});
