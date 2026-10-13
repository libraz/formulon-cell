// SpreadsheetToolbar is now a thin adapter on top of
// `Spreadsheet.mountToolbar` — the prior 5k+ LOC of React-internal ribbon UI
// (and the tests that drove it) was retired. These smoke tests
// cover the adapter's responsibilities: mounts the core toolbar against its
// host, forwards tab switches in both directions, and dispatches the
// optional review / automation / drawing callbacks as `RibbonHooks` so a
// click on the matching ribbon command lands on the host callback.

import { RIBBON_TABS, Spreadsheet } from '@libraz/formulon-cell';
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
vi.setConfig({ testTimeout: 20_000 });

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

interface Harness {
  host: HTMLElement;
  root: Root;
  unmount: () => Promise<void>;
}

const mountToolbar = async (
  _mounted: MountedReactSpreadsheet,
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
    document.body.querySelectorAll('.fc-tb__dlg').forEach((el) => {
      el.remove();
    });
    uninstallReactDomStubs();
  });

  it('mounts the core ribbon DOM into a wrapping host element', async () => {
    const onTabChange = vi.fn();
    const onToolbarReady = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="home"
        onTabChange={onTabChange}
        locale="en"
        onToolbarReady={onToolbarReady}
      />,
    );
    expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeTruthy();
    expect(onToolbarReady).toHaveBeenCalledWith(
      expect.objectContaining({ applyCommand: expect.any(Function) }),
    );
    await harness.unmount();
    expect(onToolbarReady).toHaveBeenLastCalledWith(null);
  });

  it('forwards core toolbar mount failures to onError', async () => {
    const err = new Error('toolbar mount failed');
    const onError = vi.fn();
    const onToolbarReady = vi.fn();
    const spy = vi.spyOn(Spreadsheet, 'mountToolbar').mockImplementationOnce(() => {
      throw err;
    });
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="home"
        onTabChange={() => {}}
        locale="en"
        onToolbarReady={onToolbarReady}
        onError={onError}
      />,
    );

    expect(onToolbarReady).toHaveBeenCalledWith(null);
    expect(onError).toHaveBeenCalledWith(err);
    expect(harness.host.querySelector('[data-ribbon-tab]')).toBeNull();
    await harness.unmount();
    spy.mockRestore();
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
        ribbonTabs={RIBBON_TABS}
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
    const clickCommand = (cmd: string): void => {
      harness.host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${cmd}"]`)?.click();
    };
    const clickAttr = (attr: string, value: string): void => {
      harness.host.querySelector<HTMLButtonElement>(`[data-${attr}="${value}"]`)?.click();
    };
    await act(async () => {
      clickCommand('spellingReview');
      clickCommand('accessibility');
      clickCommand('translateReview');
      // Script / AddIn ribbon buttons open a menu on plain click; the host
      // callback fires only when the user picks the action wired to its prop.
      clickCommand('script');
      clickAttr('script-action', 'custom');
      clickCommand('addIn');
      clickAttr('add-in-action', 'manage');
      clickCommand('drawPen');
      clickCommand('drawErase');
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

  describe('core default fallback', () => {
    const click = async (harness: Harness, selector: string): Promise<void> => {
      await act(async () => {
        harness.host.querySelector<HTMLButtonElement>(selector)?.click();
        await flush();
      });
    };
    const dialogs = (): number => document.body.querySelectorAll('.fc-tb__dlg').length;

    it('falls back to the core review reports when review props are omitted', async () => {
      const harness = await mountToolbar(
        mounted,
        <SpreadsheetToolbar
          instance={mounted.instance}
          activeTab="review"
          onTabChange={() => {}}
          locale="en"
        />,
      );
      for (const cmd of ['spellingReview', 'accessibility', 'translateReview']) {
        document.body.querySelectorAll('.fc-tb__dlg').forEach((el) => {
          el.remove();
        });
        await click(harness, `[data-ribbon-command="${cmd}"]`);
        expect(dialogs()).toBeGreaterThan(0);
      }
      await harness.unmount();
    });

    it('lets a provided review prop replace the core report', async () => {
      const onSpellingReview = vi.fn();
      const harness = await mountToolbar(
        mounted,
        <SpreadsheetToolbar
          instance={mounted.instance}
          activeTab="review"
          onTabChange={() => {}}
          locale="en"
          onSpellingReview={onSpellingReview}
        />,
      );
      await click(harness, '[data-ribbon-command="spellingReview"]');
      expect(onSpellingReview).toHaveBeenCalledTimes(1);
      expect(dialogs()).toBe(0);
      // Sibling review commands without a prop still use the core default.
      await click(harness, '[data-ribbon-command="accessibility"]');
      expect(dialogs()).toBeGreaterThan(0);
      await harness.unmount();
    });

    it('reaches the core script/add-in defaults when onRunScript/onAddIn are omitted', async () => {
      const harness = await mountToolbar(
        mounted,
        <SpreadsheetToolbar
          ribbonTabs={RIBBON_TABS}
          instance={mounted.instance}
          activeTab="review"
          onTabChange={() => {}}
          locale="en"
        />,
      );
      await click(harness, '[data-ribbon-command="addIn"]');
      await click(harness, '[data-add-in-action="get"]');
      expect(dialogs()).toBeGreaterThan(0);
      await harness.unmount();
    });

    it('keeps non-custom script and non-manage add-in actions on the core default', async () => {
      const onRunScript = vi.fn();
      const onAddIn = vi.fn();
      const harness = await mountToolbar(
        mounted,
        <SpreadsheetToolbar
          ribbonTabs={RIBBON_TABS}
          instance={mounted.instance}
          activeTab="review"
          onTabChange={() => {}}
          locale="en"
          onRunScript={onRunScript}
          onAddIn={onAddIn}
        />,
      );
      await click(harness, '[data-ribbon-command="addIn"]');
      await click(harness, '[data-add-in-action="get"]');
      expect(onAddIn).not.toHaveBeenCalled();
      expect(dialogs()).toBeGreaterThan(0);
      document.body.querySelectorAll('.fc-tb__dlg').forEach((el) => {
        el.remove();
      });

      await click(harness, '[data-ribbon-command="script"]');
      await click(harness, '[data-script-action="uppercase"]');
      expect(onRunScript).not.toHaveBeenCalled();
      expect(dialogs()).toBeGreaterThan(0);

      await click(harness, '[data-script-action="custom"]');
      await click(harness, '[data-add-in-action="manage"]');
      expect(onRunScript).toHaveBeenCalledTimes(1);
      expect(onAddIn).toHaveBeenCalledTimes(1);
      await harness.unmount();
    });

    it('re-mounts when a review prop appears but not when its identity changes', async () => {
      const mountSpy = vi.spyOn(Spreadsheet, 'mountToolbar');
      const render = (onSpellingReview?: () => void): ReactNode => (
        <SpreadsheetToolbar
          instance={mounted.instance}
          activeTab="review"
          onTabChange={() => {}}
          locale="en"
          onSpellingReview={onSpellingReview}
        />
      );
      const harness = await mountToolbar(mounted, render());
      const initial = mountSpy.mock.calls.length;
      const first = vi.fn();
      const second = vi.fn();
      await act(async () => {
        harness.root.render(render(first));
        await flush();
      });
      expect(mountSpy.mock.calls.length).toBe(initial + 1);
      await act(async () => {
        harness.root.render(render(second));
        await flush();
      });
      expect(mountSpy.mock.calls.length).toBe(initial + 1);
      await click(harness, '[data-ribbon-command="spellingReview"]');
      expect(first).not.toHaveBeenCalled();
      expect(second).toHaveBeenCalledTimes(1);
      await harness.unmount();
      mountSpy.mockRestore();
    });
  });

  it('routes dropdownActions overrides through core dynamic-dropdowns dispatcher', async () => {
    const onProtect = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="review"
        onTabChange={() => {}}
        locale="en"
        dropdownActions={{ applyProtectAction: onProtect }}
      />,
    );
    await act(async () => {
      harness.host.querySelector<HTMLButtonElement>('[data-protect-action="lock-cell"]')?.click();
      await flush();
    });
    expect(onProtect).toHaveBeenCalledWith('lock-cell');
    await harness.unmount();
  });

  it('preserves core Insert activation parity for PivotTable, Table, and Pictures', async () => {
    mounted.instance.setFeatures({ pivotTableDialog: true, illustrations: true });
    await flush();
    const openPivotTableDialog = vi.spyOn(mounted.instance, 'openPivotTableDialog');
    const onToolbarReady = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="insert"
        onTabChange={() => {}}
        locale="en"
        onToolbarReady={onToolbarReady}
      />,
    );
    const toolbar = onToolbarReady.mock.calls
      .map((call) => call[0])
      .find((candidate) => candidate?.applyCommand);
    expect(toolbar).toBeTruthy();
    const pivotButton = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotTableInsert"]',
    );
    const tableButton = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableInsert"]',
    );
    const pictureButton = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pictureInsert"]',
    );
    expect(pivotButton?.dataset.ribbonActivation).toBe('splitPrimary');
    expect(tableButton?.dataset.ribbonActivation).toBe('dialog');
    expect(tableButton?.dataset.ribbonMenuId).toBeUndefined();
    expect(pictureButton?.dataset.ribbonActivation).toBe('gallery');

    await act(async () => {
      expect(toolbar?.applyCommand('pivotTableInsert')).toBe(true);
      await flush();
    });
    expect(openPivotTableDialog).toHaveBeenCalledTimes(1);

    await act(async () => {
      expect(toolbar?.applyCommand('formatTableInsert')).toBe(true);
      await flush();
    });
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Create Table',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn')?.click();
    await flush();

    await act(async () => {
      pictureButton?.click();
      await flush();
    });
    const pictureMenu = harness.host.querySelector<HTMLElement>('#menu-picture-insert');
    expect(pictureMenu?.hidden).toBe(false);
    expect(pictureMenu?.classList.contains('fc-tb__menu--visual')).toBe(true);
    expect(
      pictureMenu?.querySelector<HTMLButtonElement>('[data-picture-insert="stock"]'),
    ).toBeTruthy();

    await harness.unmount();
    openPivotTableDialog.mockRestore();
  });

  it('preserves core Home split/dropdown parity for Underline and Fill', async () => {
    const onToolbarReady = vi.fn();
    const harness = await mountToolbar(
      mounted,
      <SpreadsheetToolbar
        instance={mounted.instance}
        activeTab="home"
        onTabChange={() => {}}
        locale="en"
        onToolbarReady={onToolbarReady}
      />,
    );
    const toolbar = onToolbarReady.mock.calls
      .map((call) => call[0])
      .find((candidate) => candidate?.dropdownsApi);
    expect(toolbar).toBeTruthy();

    const underline = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="underline"]',
    );
    const underlineMenu = harness.host.querySelector<HTMLDivElement>('#menu-underline');
    expect(underline?.dataset.ribbonActivation).toBe('splitToggle');
    expect(underline?.dataset.ribbonMenuId).toBe('menu-underline');
    expect(underline?.getAttribute('aria-pressed')).toBe('false');
    await act(async () => {
      underline?.click();
      await flush();
    });
    expect(underline?.getAttribute('aria-pressed')).toBe('true');
    expect(underlineMenu?.hidden).toBe(true);
    await act(async () => {
      toolbar.dropdownsApi.openDynamicRibbonDropdown(
        { command: 'underline', menuId: 'menu-underline' },
        underline as HTMLButtonElement,
      );
      await flush();
    });
    expect(underlineMenu?.hidden).toBe(false);
    expect(
      Array.from(
        harness.host.querySelectorAll<HTMLButtonElement>(
          '#menu-underline .fc-tb__menu-item--iconic',
        ),
      ).map((item) => item.textContent),
    ).toEqual(['Underline', 'Double Underline']);

    const fill = harness.host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"]');
    await act(async () => {
      fill?.click();
      await flush();
    });
    expect(
      Array.from(harness.host.querySelectorAll<HTMLButtonElement>('#menu-fill [data-fill]')).map(
        (item) => item.dataset.fill,
      ),
    ).toEqual(['down', 'right', 'up', 'left', 'group', 'series', 'justify', 'flash']);

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
