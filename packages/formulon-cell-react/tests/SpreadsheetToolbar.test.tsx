// SpreadsheetToolbar is now a thin adapter on top of
// `Spreadsheet.mountToolbar` — the prior 5k+ LOC of React-internal ribbon UI
// (and the tests that drove it) was retired in Phase 3-b. These smoke tests
// cover the adapter's responsibilities: mounts the core toolbar against its
// host, forwards tab switches in both directions, and dispatches the
// optional review / automation / drawing callbacks as `RibbonHooks` so a
// click on the matching ribbon command lands on the host callback.
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

interface Harness {
  host: HTMLElement;
  root: Root;
  unmount: () => Promise<void>;
}

const mountToolbar = async (
  mounted: MountedReactSpreadsheet,
  node: ReactNode,
): Promise<Harness> => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  let root!: Root;
  await act(async () => {
    root = createRoot(host);
    root.render(node);
    await flush();
  });
  return {
    host,
    root,
    unmount: async () => {
      await act(async () => {
        root.unmount();
        await flush();
      });
      host.remove();
    },
  };
};

describe('<SpreadsheetToolbar> (thin adapter)', () => {
  let mounted: MountedReactSpreadsheet;

  beforeEach(async () => {
    installReactDomStubs();
    mounted = await mountReactSpreadsheet({ locale: 'en' });
  });

  afterEach(async () => {
    await mounted.dispose();
    uninstallReactDomStubs();
  });

  it('mounts the core ribbon DOM into a wrapping host element', async () => {
    const onTabChange = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="home"
        onTabChange={onTabChange}
        locale="en"
      />,
    );
    expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeTruthy();
    await harness.unmount();
  });

  it('forwards tab-button clicks via onTabChange', async () => {
    const onTabChange = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="home"
        onTabChange={onTabChange}
        locale="en"
      />,
    );
    const insertTab = harness.host.querySelector<HTMLButtonElement>('[data-ribbon-tab="insert"]');
    expect(insertTab).toBeTruthy();
    await act(async () => {
      insertTab?.click();
      await flush();
    });
    expect(onTabChange).toHaveBeenCalledWith('insert');
    await harness.unmount();
  });

  it('routes ribbon review/automation/drawing commands to the matching host callback', async () => {
    const onSpellingReview = vi.fn();
    const onAccessibilityCheck = vi.fn();
    const onTranslate = vi.fn();
    const onRunScript = vi.fn();
    const onAddIn = vi.fn();
    const onDrawPen = vi.fn();
    const onDrawEraser = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="review"
        onTabChange={() => {}}
        locale="en"
        onSpellingReview={onSpellingReview}
        onAccessibilityCheck={onAccessibilityCheck}
        onTranslate={onTranslate}
        onRunScript={onRunScript}
        onAddIn={onAddIn}
        onDrawPen={onDrawPen}
        onDrawEraser={onDrawEraser}
      />,
    );
    const click = (cmd: string): void => {
      harness.host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${cmd}"]`)?.click();
    };
    await act(async () => {
      click('spellingReview');
      click('accessibility');
      click('translateReview');
      click('script');
      click('addIn');
      click('drawPen');
      click('drawErase');
      await flush();
    });
    expect(onSpellingReview).toHaveBeenCalledTimes(1);
    expect(onAccessibilityCheck).toHaveBeenCalledTimes(1);
    expect(onTranslate).toHaveBeenCalledTimes(1);
    expect(onRunScript).toHaveBeenCalledTimes(1);
    expect(onAddIn).toHaveBeenCalledTimes(1);
    expect(onDrawPen).toHaveBeenCalledTimes(1);
    expect(onDrawEraser).toHaveBeenCalledTimes(1);
    await harness.unmount();
  });

  it('reacts to external activeTab prop changes without re-mounting the core toolbar', async () => {
    const onTabChange = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="home"
        onTabChange={onTabChange}
        locale="en"
      />,
    );
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeTruthy();
    await act(async () => {
      harness.root.render(
        <SpreadsheetToolbar
          instance={mounted.instance}
          activeTab="data"
          onTabChange={onTabChange}
          locale="en"
        />,
      );
      await flush();
    });
    expect(harness.host.querySelector('[data-ribbon-panel="data"]:not([hidden])')).toBeTruthy();
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeFalsy();
    await harness.unmount();
  });

  it('cleans up the core toolbar when React unmounts the component', async () => {
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="home"
        onTabChange={() => {}}
        locale="en"
      />,
    );
    expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    await harness.unmount();
    // The wrapper host element is detached, so its descendants are gone too.
    expect(document.body.contains(harness.host)).toBe(false);
  });
});
