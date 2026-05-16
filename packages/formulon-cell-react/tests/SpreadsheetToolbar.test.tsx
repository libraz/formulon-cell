import {
  applyCellStyle,
  applyValueFilter,
  buildRibbonModel,
  type ConditionalRule,
  commentAt,
  type DeepPartial,
  EMPTY_ACTIVE_STATE,
  formatAsTable,
  listComments,
  mutators,
  projectActiveState,
  RIBBON_TAB_LABELS,
  type RibbonTab,
  type Strings,
  setComment,
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

const dateSerial = (year: number, month: number, day: number): number =>
  Date.UTC(year, month - 1, day) / 86_400_000 + 25569;

interface ToolbarHarness {
  host: HTMLElement;
  root: Root;
  rerender(
    activeTab: RibbonTab,
    onTabChange?: (tab: RibbonTab) => void,
    fileHandlers?: FileHandlers,
  ): Promise<void>;
  unmount(): Promise<void>;
}

interface FileHandlers {
  onNewWorkbook?: () => void;
  onOpenWorkbook?: () => void;
  onSaveWorkbook?: () => void;
  onSaveWorkbookAs?: () => void;
}

async function renderToolbar(
  mounted: MountedReactSpreadsheet | null,
  initial: {
    activeTab: RibbonTab;
    onTabChange: (tab: RibbonTab) => void;
    locale?: string;
    fileHandlers?: FileHandlers;
  },
): Promise<ToolbarHarness> {
  installReactDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);

  let lastFileHandlers = initial.fileHandlers ?? {};

  const render = (
    activeTab: RibbonTab,
    onTabChange: (tab: RibbonTab) => void,
    fileHandlers: FileHandlers,
  ): ReactNode => (
    <SpreadsheetToolbar
      instance={mounted?.instance ?? null}
      activeTab={activeTab}
      onTabChange={onTabChange}
      locale={initial.locale ?? 'en'}
      {...fileHandlers}
    />
  );

  await act(async () => {
    root.render(render(initial.activeTab, initial.onTabChange, lastFileHandlers));
    await flush();
  });

  let lastTab = initial.activeTab;
  let lastChange = initial.onTabChange;

  return {
    host,
    root,
    async rerender(activeTab, onTabChange, fileHandlers) {
      lastTab = activeTab;
      if (onTabChange) lastChange = onTabChange;
      if (fileHandlers) lastFileHandlers = fileHandlers;
      await act(async () => {
        root.render(render(lastTab, lastChange, lastFileHandlers));
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

  it('renders a tab button for every ribbon tab using RIBBON_TAB_LABELS', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const tabButtons = toolbar.host.querySelectorAll<HTMLButtonElement>('[role="tab"]');
    const labels = Array.from(tabButtons).map((b) => b.textContent?.trim());
    const expected = (Object.keys(RIBBON_TAB_LABELS) as RibbonTab[]).map(
      (id) => RIBBON_TAB_LABELS[id].en,
    );
    expect(labels).toEqual(expected);
    expect(tabButtons[0]?.className).toContain('demo__ribbon-tab--file');
    expect(toolbar.host.querySelector<HTMLElement>('.demo__ribbon')?.className).toContain(
      'demo__ribbon--office365-home',
    );

    await toolbar.rerender('data');
    expect(toolbar.host.querySelector<HTMLElement>('.demo__ribbon')?.className).not.toContain(
      'demo__ribbon--office365-home',
    );
  });

  it('exposes the shared core ribbon commands as DOM command ids', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'file', onTabChange: vi.fn() });
    const exposed = new Set<string>();

    for (const tab of buildRibbonModel('en')) {
      await toolbar.rerender(tab.id);
      for (const el of toolbar.host.querySelectorAll<HTMLElement>('[data-ribbon-command]')) {
        const command = el.dataset.ribbonCommand;
        if (command) exposed.add(command);
      }
    }

    const coreIds = new Set(
      buildRibbonModel('en')
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .map((command) => command.id),
    );
    const missing = Array.from(coreIds).filter((id) => !exposed.has(id));

    expect(missing).toEqual([]);
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

  it('keeps host-callback ribbon buttons enabled without a spreadsheet instance', async () => {
    installReactDomStubs();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    const onTabChange = vi.fn();
    const onTranslate = vi.fn();
    const onRunScript = vi.fn();
    const onAddIn = vi.fn();

    const render = async (activeTab: RibbonTab): Promise<void> => {
      await act(async () => {
        root.render(
          <SpreadsheetToolbar
            instance={null}
            activeTab={activeTab}
            onTabChange={onTabChange}
            locale="en"
            onTranslate={onTranslate}
            onRunScript={onRunScript}
            onAddIn={onAddIn}
          />,
        );
        await flush();
      });
    };

    try {
      await render('review');
      const translate = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find(
        (button) => button.getAttribute('aria-label') === 'Translate',
      );
      expect(translate?.disabled).toBe(false);
      await act(async () => {
        translate?.click();
        await flush();
      });
      expect(onTranslate).toHaveBeenCalledTimes(1);

      await render('automate');
      const script = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find(
        (button) => button.getAttribute('aria-label') === 'Script',
      );
      expect(script?.disabled).toBe(false);
      await act(async () => {
        script?.click();
        await flush();
      });
      expect(onRunScript).toHaveBeenCalledTimes(1);

      await render('acrobat');
      const addIn = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find(
        (button) => button.getAttribute('aria-label') === 'Add-ins',
      );
      expect(addIn?.disabled).toBe(false);
      await act(async () => {
        addIn?.click();
        await flush();
      });
      const myAddIn = host.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="addIn"] [data-cell-action="my"]',
      );
      await act(async () => {
        myAddIn?.click();
        await flush();
      });
      expect(host.textContent).toContain('My Add-ins');
      expect(host.textContent).toContain('Built-in add-ins');
      expect(host.textContent).toContain('External add-ins');
      expect(onAddIn).not.toHaveBeenCalled();
    } finally {
      await act(async () => {
        root.unmount();
        await flush();
      });
      host.remove();
      uninstallReactDomStubs();
    }
  });

  it('routes Acrobat tab PDF and built-in Add-ins menu actions', async () => {
    mounted = await mountReactSpreadsheet();
    const print = vi.spyOn(mounted.instance, 'print').mockImplementation(() => undefined);
    toolbar = await renderToolbar(mounted, { activeTab: 'acrobat', onTabChange: vi.fn() });

    const addIn = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="addIn"] button',
    );
    expect(addIn?.disabled).toBe(false);
    await act(async () => {
      addIn?.click();
      await flush();
    });
    const getAddIn = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="addIn"] [data-cell-action="get"]',
    );
    await act(async () => {
      getAddIn?.click();
      await flush();
    });
    expect(toolbar.host.textContent).toContain('Office Add-ins store');

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('.demo__modal-x')?.click();
      await flush();
    });
    await act(async () => {
      addIn?.click();
      await flush();
    });
    const myAddIn = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="addIn"] [data-cell-action="my"]',
    );
    await act(async () => {
      myAddIn?.click();
      await flush();
    });
    expect(toolbar.host.textContent).toContain('My Add-ins');
    expect(toolbar.host.textContent).toContain('Built-in add-ins');

    const pdf = toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="pdf"] button');
    expect(pdf?.disabled).toBe(false);
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('.demo__modal-x')?.click();
      await flush();
    });
    await act(async () => {
      pdf?.click();
      await flush();
    });
    const createPdf = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pdf"] [data-cell-action="create"]',
    );
    await act(async () => {
      createPdf?.click();
      await flush();
    });

    expect(print).toHaveBeenCalledWith('pdf');
    expect(toolbar.host.textContent).toContain('Create PDF');
    expect(toolbar.host.textContent).toContain('PDF export has been sent');

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('.demo__modal-x')?.click();
      await flush();
    });
    await act(async () => {
      pdf?.click();
      await flush();
    });
    const sharePdf = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pdf"] [data-cell-action="share"]',
    );
    await act(async () => {
      sharePdf?.click();
      await flush();
    });
    expect(print).toHaveBeenLastCalledWith('pdf');
    expect(toolbar.host.textContent).toContain('Create PDF and Share Link');
    expect(toolbar.host.textContent).toContain('PDF export is ready.');
  });

  it('localizes React ribbon command titles and accessibility labels in Japanese', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, {
      activeTab: 'home',
      onTabChange: vi.fn(),
      locale: 'ja',
    });

    const ariaLabels = (): string[] => {
      if (!toolbar) return [];
      return Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).map(
        (button) => button.getAttribute('aria-label') ?? '',
      );
    };

    expect(ariaLabels()).toEqual(expect.arrayContaining(['太字 (⌘B)', '数値', '上揃え']));

    await toolbar.rerender('insert');
    expect(ariaLabels()).toEqual(
      expect.arrayContaining(['ピボットテーブル', '名前', 'メモを挿入']),
    );

    await toolbar.rerender('draw');
    expect(ariaLabels()).toContain('消しゴム');

    await toolbar.rerender('formulas');
    expect(ariaLabels()).toEqual(
      expect.arrayContaining(['関数の挿入', 'SUM の引数', '参照元', '再計算 (F9)']),
    );

    await toolbar.rerender('data');
    expect(ariaLabels()).toEqual(expect.arrayContaining(['昇順で並べ替え', 'リンク']));

    await toolbar.rerender('pageLayout');
    expect(ariaLabels()).toContain('ページ設定');

    await toolbar.rerender('view');
    expect(ariaLabels()).toEqual(expect.arrayContaining(['ウィンドウ枠', 'ズーム 100%', '保護']));
  });

  it('updates React ribbon copy from live instance i18n overrides', async () => {
    mounted = await mountReactSpreadsheet({ locale: 'en' });
    toolbar = await renderToolbar(mounted, {
      activeTab: 'home',
      onTabChange: vi.fn(),
      locale: 'en',
    });
    const currentToolbar = toolbar;
    const currentMounted = mounted;

    const boldButton = (): HTMLButtonElement | undefined =>
      Array.from(currentToolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
        (button) => button.dataset.ribbonCommand === 'bold',
      );

    expect(boldButton()?.getAttribute('aria-label')).toBe('Bold (⌘B)');

    await act(async () => {
      currentMounted.instance.i18n.extend('en', {
        ribbon: { bold: 'Strong' },
      } as DeepPartial<Strings>);
      await flush();
    });

    expect(boldButton()?.getAttribute('aria-label')).toBe('Strong (⌘B)');
  });

  it('uses live i18n overrides when toolbar and instance locale tags share a language', async () => {
    mounted = await mountReactSpreadsheet({ locale: 'ja-JP' });
    toolbar = await renderToolbar(mounted, {
      activeTab: 'home',
      onTabChange: vi.fn(),
      locale: 'ja',
    });
    const currentToolbar = toolbar;
    const currentMounted = mounted;

    const boldButton = (): HTMLButtonElement | undefined =>
      Array.from(currentToolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
        (button) => button.dataset.ribbonCommand === 'bold',
      );

    expect(boldButton()?.getAttribute('aria-label')).toBe('太字 (⌘B)');

    await act(async () => {
      currentMounted.instance.i18n.extend('ja-JP', {
        ribbon: { bold: '強調' },
      } as DeepPartial<Strings>);
      await flush();
    });

    expect(boldButton()?.getAttribute('aria-label')).toBe('強調 (⌘B)');
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

  it('collapses and restores the command ribbon with global Ctrl+F1 and tab double-click', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const shell = toolbar.host.querySelector<HTMLElement>('.demo__ribbon-shell');
    const tablist = toolbar.host.querySelector<HTMLElement>('.demo__ribbon-tabs');
    const homeTab = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('[role="tab"]'),
    ).find((t) => t.textContent?.includes('Home'));
    const collapseButton = toolbar.host.querySelector<HTMLButtonElement>('.demo__ribbon-toggle');
    expect(shell?.className).not.toContain('demo__ribbon-shell--collapsed');
    expect(tablist?.dataset.ribbonCollapsed).toBe('false');
    expect(collapseButton?.getAttribute('aria-label')).toBe('Ribbon Display Options');
    expect(collapseButton?.getAttribute('aria-expanded')).toBe('false');

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true }));
      await flush();
    });

    expect(shell?.className).toContain('demo__ribbon-shell--collapsed');
    expect(tablist?.dataset.ribbonCollapsed).toBe('true');

    await act(async () => {
      collapseButton?.click();
      await flush();
    });

    const collapsedOption = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('.demo__ribbon-display-option'),
    ).find((b) => b.textContent === 'Show tabs only');
    expect(collapseButton?.getAttribute('aria-expanded')).toBe('true');
    expect(collapsedOption?.getAttribute('aria-checked')).toBe('true');

    await act(async () => {
      document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
      await flush();
    });

    expect(toolbar.host.querySelector('.demo__ribbon-display-menu')).toBeNull();

    await act(async () => {
      collapseButton?.click();
      await flush();
    });

    const reopenedExpandedOption = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('.demo__ribbon-display-option'),
    ).find((b) => b.textContent === 'Always show Ribbon');

    await act(async () => {
      reopenedExpandedOption?.click();
      await flush();
    });

    expect(shell?.className).not.toContain('demo__ribbon-shell--collapsed');
    expect(tablist?.dataset.ribbonCollapsed).toBe('false');

    await act(async () => {
      collapseButton?.focus();
      collapseButton?.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }),
      );
      await flush();
    });

    const keyboardExpandedOption = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('.demo__ribbon-display-option'),
    ).find((b) => b.textContent === 'Always show Ribbon');
    const keyboardCollapsedOption = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('.demo__ribbon-display-option'),
    ).find((b) => b.textContent === 'Show tabs only');
    expect(document.activeElement).toBe(keyboardExpandedOption);

    await act(async () => {
      keyboardExpandedOption?.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'End', bubbles: true }),
      );
      await flush();
    });

    expect(document.activeElement).toBe(keyboardCollapsedOption);

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
      await flush();
    });

    expect(toolbar.host.querySelector('.demo__ribbon-display-menu')).toBeNull();

    await act(async () => {
      homeTab?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      await flush();
    });

    expect(shell?.className).toContain('demo__ribbon-shell--collapsed');
    expect(tablist?.dataset.ribbonCollapsed).toBe('true');
  });

  it('renders File as an Excel-style backstage view', async () => {
    mounted = await mountReactSpreadsheet();
    const onTabChange = vi.fn();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange });

    await toolbar.rerender('file', onTabChange);

    const backstage = toolbar.host.querySelector<HTMLElement>('.demo__backstage');
    expect(backstage).toBeDefined();
    expect(backstage?.getAttribute('role')).toBe('dialog');
    expect(backstage?.querySelector('.demo__backstage-navitem--active')?.textContent).toBe('Info');
    expect(backstage?.textContent).toContain('Workbook Information');
    expect(backstage?.textContent).toContain('Properties');
    expect(backstage?.textContent).toContain('Inspect Workbook');
    expect(backstage?.textContent).toContain('Open');
    expect(backstage?.textContent).toContain('Save As');
    expect(backstage?.textContent).toContain('Print');
    expect(
      backstage
        ?.querySelector<HTMLButtonElement>('.demo__backstage-navitem')
        ?.getAttribute('aria-label'),
    ).toBe('Back');
    expect(toolbar.host.querySelector('.demo__ribbon')).toBeNull();
    expect(
      Array.from(
        backstage?.querySelectorAll<HTMLButtonElement>('.demo__backstage-card') ?? [],
      ).find((b) => b.textContent?.includes('New'))?.disabled,
    ).toBe(true);
    const cards = Array.from(
      backstage?.querySelectorAll<HTMLButtonElement>('.demo__backstage-card') ?? [],
    );
    expect(cards).toHaveLength(6);
    expect(cards.map((b) => b.querySelector('strong')?.textContent)).toEqual([
      'New',
      'Open',
      'Save',
      'Save As',
      'Print',
      'Options',
    ]);
    expect(cards.some((b) => b.textContent?.includes('Format Cells'))).toBe(false);
    expect(cards.find((b) => b.textContent?.includes('Options'))?.dataset.ribbonCommand).toBe(
      'pageSetup',
    );

    const onNewWorkbook = vi.fn();
    const onOpenWorkbook = vi.fn();
    const onSaveWorkbook = vi.fn();
    const onSaveWorkbookAs = vi.fn();
    await toolbar.rerender('file', onTabChange, {
      onNewWorkbook,
      onOpenWorkbook,
      onSaveWorkbook,
      onSaveWorkbookAs,
    });

    const nav = toolbar.host.querySelector<HTMLElement>('.demo__backstage-nav');
    const newButton = Array.from(nav?.querySelectorAll<HTMLButtonElement>('button') ?? []).find(
      (b) => b.textContent === 'New',
    );
    expect(newButton?.disabled).toBe(false);
    const manageButton = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('.demo__backstage-command'),
    ).find((b) => b.textContent?.includes('Manage Workbook'));
    expect(manageButton?.disabled).toBe(false);
    const protectButton = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__backstage-command[data-ribbon-command="protect"]',
    );
    const inspectButton = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__backstage-command[data-ribbon-command="inspect"]',
    );
    expect(protectButton?.textContent).toContain('Protect Workbook');
    expect(inspectButton?.textContent).toContain('Inspect Workbook');

    await act(async () => {
      newButton?.click();
      await flush();
    });

    expect(onNewWorkbook).toHaveBeenCalledTimes(1);

    await act(async () => {
      protectButton?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().protection.workbookStructure).toBeDefined();
    expect(
      toolbar.host
        .querySelector<HTMLButtonElement>('.demo__backstage-command[data-ribbon-command="protect"]')
        ?.getAttribute('aria-pressed'),
    ).toBe('true');

    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('.demo__backstage-command[data-ribbon-command="protect"]')
        ?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().protection.workbookStructure).toBeUndefined();

    await act(async () => {
      inspectButton?.click();
      await flush();
    });

    expect(toolbar.host.querySelector('.demo__modal')?.textContent).toContain(
      'Spreadsheet compatibility',
    );

    await act(async () => {
      manageButton?.click();
      await flush();
    });

    expect(onSaveWorkbookAs).toHaveBeenCalledTimes(1);

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
      await flush();
    });

    expect(onTabChange).toHaveBeenCalledWith('home');
  });

  it('localizes Backstage workbook inspection results', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, {
      activeTab: 'file',
      onTabChange: vi.fn(),
      locale: 'ja',
    });

    const inspectButton = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__backstage-command[data-ribbon-command="inspect"]',
    );
    await act(async () => {
      inspectButton?.click();
      await flush();
    });

    const reportText = toolbar.host.querySelector('.demo__modal')?.textContent ?? '';
    expect(reportText).toContain('スプレッドシート互換');
    expect(reportText).toContain('書き込み可');
    expect(reportText).toContain('セルの書式');
    expect(reportText).toContain('エンジンが対応している場合');
    expect(reportText).not.toContain('Formatting can be shown');
  });

  it('exposes Excel-style shortcut hints on routed ribbon buttons', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const formatCells = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"] button',
    );
    expect(formatCells?.getAttribute('aria-keyshortcuts')).toContain('Control+1');

    const find = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find((b) =>
      b.getAttribute('aria-label')?.startsWith('Find'),
    );
    expect(find?.getAttribute('aria-keyshortcuts')).toContain('Control+F');

    await toolbar.rerender('formulas');
    const namedRanges = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (b) => b.getAttribute('aria-label') === 'Name manager',
    );
    expect(namedRanges?.getAttribute('aria-keyshortcuts')).toBe('Control+F3');

    const fx = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (b) => b.getAttribute('aria-label') === 'Insert function',
    );
    expect(fx?.getAttribute('aria-keyshortcuts')).toBe('Shift+F3');
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

  it('routes Home clipboard paste and format painter buttons to host actions', async () => {
    mounted = await mountReactSpreadsheet();
    const originalExecCommand = document.execCommand;
    const execCommand = vi.fn(() => true);
    Object.defineProperty(document, 'execCommand', {
      configurable: true,
      value: execCommand,
    });
    const formatPainter = vi.spyOn(mounted.instance.formatPainter, 'activate');
    const pasteSpecialApply = vi.fn(() => true);
    mounted.instance.pasteSpecial = pasteSpecialApply;
    const openInsertCopiedCells = vi.fn();
    mounted.instance.openInsertCopiedCells = openInsertCopiedCells;
    const pasteSpecial = vi.spyOn(mounted.instance, 'openPasteSpecial');
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    try {
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"] button')?.click();
        await flush();
      });
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-cell-action="paste"]')?.click();
        await flush();
      });
      expect(execCommand).toHaveBeenCalledWith('paste');

      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"] button')?.click();
        await flush();
      });
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-cell-action="pasteValues"]')?.click();
        await flush();
      });
      expect(pasteSpecialApply).toHaveBeenCalledWith({
        what: 'values',
        operation: 'none',
        skipBlanks: false,
        transpose: false,
      });

      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"] button')?.click();
        await flush();
      });
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-cell-action="pasteTranspose"]')?.click();
        await flush();
      });
      expect(pasteSpecialApply).toHaveBeenCalledWith({
        what: 'all',
        operation: 'none',
        skipBlanks: false,
        transpose: true,
      });

      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"] button')?.click();
        await flush();
      });
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-cell-action="insertCopiedCells"]')?.click();
        await flush();
      });
      expect(openInsertCopiedCells).toHaveBeenCalledTimes(1);

      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"] button')?.click();
        await flush();
      });
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>('[data-cell-action="pasteSpecial"]')?.click();
        await flush();
      });
      expect(pasteSpecial).toHaveBeenCalledTimes(1);

      await act(async () => {
        toolbar?.host
          .querySelector<HTMLButtonElement>('[data-ribbon-command="formatPainter"]')
          ?.click();
        await flush();
      });
      expect(formatPainter).toHaveBeenCalledWith(false);
    } finally {
      Object.defineProperty(document, 'execCommand', {
        configurable: true,
        value: originalExecCommand,
      });
    }
  });

  it('applies Home > Alignment indent commands to the selected range with undo support', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const indentIncrease = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="indentIncrease"]',
    );
    const indentDecrease = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="indentDecrease"]',
    );
    expect(indentIncrease?.disabled).toBe(false);
    expect(indentDecrease?.disabled).toBe(false);

    await act(async () => {
      indentIncrease?.click();
      await flush();
    });
    for (const key of ['0:0:0', '0:0:1', '0:1:0', '0:1:1']) {
      expect(mounted.instance.store.getState().format.formats.get(key)?.indent).toBe(1);
    }

    await act(async () => {
      indentDecrease?.click();
      await flush();
    });
    for (const key of ['0:0:0', '0:0:1', '0:1:0', '0:1:1']) {
      expect(mounted.instance.store.getState().format.formats.get(key)?.indent).toBe(0);
    }

    await act(async () => {
      mounted.instance.undo();
      await flush();
    });
    for (const key of ['0:0:0', '0:0:1', '0:1:0', '0:1:1']) {
      expect(mounted.instance.store.getState().format.formats.get(key)?.indent).toBe(1);
    }

    await act(async () => {
      mounted.instance.redo();
      await flush();
    });
    for (const key of ['0:0:0', '0:0:1', '0:1:0', '0:1:1']) {
      expect(mounted.instance.store.getState().format.formats.get(key)?.indent).toBe(0);
    }
  });

  it('applies Home font, alignment, and number-format commands to the selected range', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const click = async (command: string): Promise<void> => {
      await act(async () => {
        toolbar?.host
          .querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`)
          ?.click();
        await flush();
      });
    };

    await click('fontGrow');
    await click('italic');
    await click('underline');
    await click('strike');
    await click('middle');
    await click('alignC');
    await click('currency');
    await click('decDown');

    for (const key of ['0:0:0', '0:0:1']) {
      const fmt = mounted.instance.store.getState().format.formats.get(key);
      expect(fmt).toMatchObject({
        fontSize: 12,
        italic: true,
        underline: true,
        strike: true,
        vAlign: 'middle',
        align: 'center',
        numFmt: { kind: 'currency', decimals: 1, symbol: '$' },
      });
    }

    await click('fontShrink');
    await click('bottomAlign');
    await click('alignR');
    await click('decUp');

    for (const key of ['0:0:0', '0:0:1']) {
      const fmt = mounted.instance.store.getState().format.formats.get(key);
      expect(fmt).toMatchObject({
        fontSize: 11,
        vAlign: 'bottom',
        align: 'right',
        numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      });
    }

    await click('alignL');
    for (const key of ['0:0:0', '0:0:1']) {
      expect(mounted.instance.store.getState().format.formats.get(key)?.align).toBe('left');
    }
  });

  it('opens the text-orientation menu and applies rotation presets', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const orientationButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textOrientation"] button',
    );
    expect(orientationButton).toBeDefined();

    await act(async () => {
      orientationButton?.click();
      await flush();
    });

    expect(toolbar.host.querySelector('[data-cell-action="angleCounterclockwise"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cell-action="rotateTextDown"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cell-action="horizontalText"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cell-action="formatAlignment"]')).toBeTruthy();

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="rotateTextDown"]')?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().format.formats.get('0:0:0')?.rotation).toBe(-90);

    await toolbar.unmount();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });
    const activeOrientationButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textOrientation"] button',
    );
    expect(activeOrientationButton?.className).toContain('demo__rb--active');
    await act(async () => {
      activeOrientationButton?.click();
      await flush();
    });
    expect(
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="rotateTextDown"]')
        ?.className,
    ).toContain('demo__rb--active');

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="horizontalText"]')?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().format.formats.get('0:0:0')?.rotation).toBe(0);

    const openFormatDialog = vi.spyOn(mounted.instance, 'openFormatDialog');
    await act(async () => {
      activeOrientationButton?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-cell-action="formatAlignment"]')
        ?.click();
      await flush();
    });

    expect(openFormatDialog).toHaveBeenCalledWith('align');
  });

  it('marks the Merge menu active for merged and merge-centered cells', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.mergeRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const mergeButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="merge"] button',
    );
    expect(mergeButton?.className).toContain('demo__rb--active');

    await act(async () => {
      mergeButton?.click();
      await flush();
    });
    expect(
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"] [aria-checked="true"]')
        ?.textContent,
    ).toContain('Merge cells');

    mutators.setCellFormat(mounted.instance.store, { sheet: 0, row: 0, col: 0 }, { align: 'center' });
    await toolbar.unmount();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"] button')?.click();
      await flush();
    });
    expect(
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"] [aria-checked="true"]')
        ?.textContent,
    ).toContain('Merge & Center');
  });

  it('opens Format Cells on the Number tab from More Number Formats', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });
    const openFormatDialog = vi.spyOn(mounted.instance, 'openFormatDialog');

    const numberButton = toolbar.host.querySelector<HTMLButtonElement>(
      'button[aria-label="Number"]',
    );
    await act(async () => {
      numberButton?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-fc-value="more"]')?.click();
      await flush();
    });

    expect(openFormatDialog).toHaveBeenCalledWith('number');
  });

  it('applies comma style as a fixed number format with thousands grouping', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="comma"]')?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().format.formats.get('0:0:0')?.numFmt).toEqual({
      kind: 'fixed',
      decimals: 2,
      thousands: true,
    });
  });

  it('applies the selected border line style through the border preset menu', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const lineStyleButton = toolbar.host.querySelector<HTMLButtonElement>(
      'button[aria-label="Border line style"]',
    );
    expect(lineStyleButton).toBeDefined();

    await act(async () => {
      lineStyleButton?.click();
      await flush();
    });

    expect(
      Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('[role="option"]')).map(
        (option) => option.dataset.fcValue,
      ),
    ).toEqual(
      expect.arrayContaining([
        'hair',
        'mediumDashed',
        'dashDot',
        'mediumDashDot',
        'dashDotDot',
        'mediumDashDotDot',
        'slantDashDot',
      ]),
    );

    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[role="option"][data-fc-value="thick"]')
        ?.click();
      await flush();
    });

    const presetButton = toolbar.host.querySelector<HTMLButtonElement>(
      'button[aria-label="Border pattern"]',
    );
    expect(presetButton).toBeDefined();

    await act(async () => {
      presetButton?.click();
      await flush();
    });

    expect(
      Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('[role="option"]')).map(
        (option) => option.dataset.fcValue,
      ),
    ).toEqual(
      expect.arrayContaining([
        'thickOutline',
        'inside',
        'insideHorizontal',
        'insideVertical',
        'thickBottom',
        'topAndBottom',
        'topAndThickBottom',
        'topAndDoubleBottom',
        'diagonalDown',
        'diagonalUp',
      ]),
    );

    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[role="option"][data-fc-value="outline"]')
        ?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().format.formats.get('0:0:0')?.borders).toEqual({
      top: { style: 'thick', color: '#000000' },
      right: { style: 'thick', color: '#000000' },
      bottom: { style: 'thick', color: '#000000' },
      left: { style: 'thick', color: '#000000' },
    });

    const openFormatDialog = vi.spyOn(mounted.instance, 'openFormatDialog');
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="moreBorders"]')?.click();
      await flush();
    });

    expect(openFormatDialog).toHaveBeenCalledWith('border');

    const borderDraw = mounted.instance.borderDraw;
    expect(borderDraw).toBeDefined();
    if (!borderDraw) throw new Error('borderDraw extension is not available');
    const setStyle = vi.spyOn(borderDraw, 'setStyle');
    const setColor = vi.spyOn(borderDraw, 'setColor');

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="drawBorder"]')?.click();
      await flush();
    });
    expect(mounted.instance.borderDraw?.getMode()).toBe('draw');

    await act(async () => {
      lineStyleButton?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[role="option"][data-fc-value="dashed"]')
        ?.click();
      await flush();
    });
    expect(setStyle).toHaveBeenCalledWith('dashed');

    const borderColorInput = toolbar.host.querySelector<HTMLInputElement>(
      '[data-ribbon-command="borderColor"] input[type="color"]',
    );
    await act(async () => {
      if (borderColorInput) {
        const valueSetter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          'value',
        )?.set;
        valueSetter?.call(borderColorInput, '#c00000');
        borderColorInput.dispatchEvent(new Event('input', { bubbles: true }));
        borderColorInput.dispatchEvent(new Event('change', { bubbles: true }));
      }
      await flush();
    });
    expect(setColor).toHaveBeenCalledWith('#c00000');

    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="drawBorderGrid"]')
        ?.click();
      await flush();
    });
    expect(mounted.instance.borderDraw?.getMode()).toBe('grid');

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="eraseBorder"]')?.click();
      await flush();
    });
    expect(mounted.instance.borderDraw?.getMode()).toBe('erase');
  });

  it('marks View > Formula Bar as an active visibility toggle', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    const formulaBarButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="viewFormulaBar"]',
    );
    expect(formulaBarButton?.className).toContain('demo__rb--active');

    await act(async () => {
      formulaBarButton?.click();
      await flush();
    });

    const refreshed = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="viewFormulaBar"]',
    );
    expect(refreshed?.className).not.toContain('demo__rb--active');
    expect(mounted.instance.host.querySelector('.fc-host__formulabar')).toBeNull();
  });

  it('toggles View > Show flags for gridlines, headings, formulas, and R1C1', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    const click = async (command: string): Promise<void> => {
      await act(async () => {
        toolbar?.host
          .querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`)
          ?.click();
        await flush();
      });
    };

    await click('viewGridlines');
    expect(mounted.instance.store.getState().ui.showGridLines).toBe(false);

    await click('viewHeadings');
    expect(mounted.instance.store.getState().ui.showHeaders).toBe(false);

    await click('viewFormulas');
    expect(mounted.instance.store.getState().ui.showFormulas).toBe(true);

    await click('viewR1C1');
    expect(mounted.instance.store.getState().ui.r1c1).toBe(true);

    expect(
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="viewGridlines"]')
        ?.className,
    ).not.toContain('demo__rb--active');
    expect(
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="viewR1C1"]')?.className,
    ).toContain('demo__rb--active');
  });

  it('switches View > Workbook Views modes and marks the active mode', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    const normal = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="viewNormal"]',
    );
    const pageLayout = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="viewPageLayout"]',
    );
    const pageBreak = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="viewPageBreakPreview"]',
    );
    expect(normal?.className).toContain('demo__rb--active');

    await act(async () => {
      pageBreak?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().ui.workbookView).toBe('pageBreakPreview');
    expect(mounted.instance.host.dataset.fcWorkbookView).toBe('pageBreakPreview');
    expect(pageLayout?.className).not.toContain('demo__rb--active');
    expect(pageBreak?.className).toContain('demo__rb--active');
  });

  it('applies View > Window > Freeze Panes menu actions with undo support', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 3, col: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    const pickFreeze = async (action: string): Promise<void> => {
      await act(async () => {
        toolbar?.host
          .querySelector<HTMLButtonElement>('[data-ribbon-command="freeze"] button')
          ?.click();
        await flush();
      });
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>(`[data-cell-action="${action}"]`)?.click();
        await flush();
      });
    };

    await pickFreeze('topRow');
    expect(mounted.instance.store.getState().layout).toMatchObject({
      freezeRows: 1,
      freezeCols: 0,
    });
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().layout).toMatchObject({
      freezeRows: 0,
      freezeCols: 0,
    });

    await pickFreeze('firstColumn');
    expect(mounted.instance.store.getState().layout).toMatchObject({
      freezeRows: 0,
      freezeCols: 1,
    });

    await pickFreeze('panes');
    expect(mounted.instance.store.getState().layout).toMatchObject({
      freezeRows: 3,
      freezeCols: 2,
    });

    await pickFreeze('none');
    expect(mounted.instance.store.getState().layout).toMatchObject({
      freezeRows: 0,
      freezeCols: 0,
    });
  });

  it('hides and unhides selected rows and columns from View > Window', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 2, c0: 1, r1: 3, c1: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    const pickWindow = async (action: string): Promise<void> => {
      await act(async () => {
        toolbar?.host
          .querySelector<HTMLButtonElement>('[data-ribbon-command="windowVisibility"] button')
          ?.click();
        await flush();
      });
      await act(async () => {
        toolbar?.host.querySelector<HTMLButtonElement>(`[data-cell-action="${action}"]`)?.click();
        await flush();
      });
    };

    await pickWindow('hideRows');
    expect(mounted.instance.store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(mounted.instance.store.getState().layout.hiddenRows.has(3)).toBe(true);

    await pickWindow('showRows');
    expect(mounted.instance.store.getState().layout.hiddenRows.has(2)).toBe(false);
    expect(mounted.instance.store.getState().layout.hiddenRows.has(3)).toBe(false);

    await pickWindow('hideCols');
    expect(mounted.instance.store.getState().layout.hiddenCols.has(1)).toBe(true);
    expect(mounted.instance.store.getState().layout.hiddenCols.has(2)).toBe(true);

    await pickWindow('showCols');
    expect(mounted.instance.store.getState().layout.hiddenCols.has(1)).toBe(false);
    expect(mounted.instance.store.getState().layout.hiddenCols.has(2)).toBe(false);
  });

  it('zooms View > Zoom > Zoom to Selection to fit the selected range', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setViewportSize(mounted.instance.store, 20, 16);
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 9, c1: 7 });
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    await act(async () => {
      toolbar?.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="zoomSelection"]')
        ?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().viewport.zoom).toBe(2);
  });

  it('applies View > Zoom fixed percentage commands', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    await act(async () => {
      toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="zoom75"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().viewport.zoom).toBe(0.75);

    await act(async () => {
      toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="zoom125"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().viewport.zoom).toBe(1.25);
  });

  it('applies a custom View > Zoom dialog percentage', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    await act(async () => {
      toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="zoomDialog"]')?.click();
      await flush();
    });
    const input = toolbar.host.querySelector<HTMLInputElement>('.demo__modal input[type="number"]');
    await act(async () => {
      if (input) {
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
          input,
          '150',
        );
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }
      await flush();
    });
    const okButton = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('.demo__btn'),
    ).find((button) => button.textContent === 'OK');
    await act(async () => {
      okButton?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().viewport.zoom).toBe(1.5);
  });

  it('saves, activates, and deletes a sheet view from the View ribbon', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'view', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="sheetViewSave"]')?.click();
      await flush();
    });
    await toolbar.rerender('view');

    let state = mounted.instance.store.getState();
    expect(state.sheetViews.views).toHaveLength(1);
    const saved = state.sheetViews.views[0];
    expect(saved?.name).toBe('Views 1');
    expect(state.sheetViews.activeViewId).toBe(saved?.id);

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="sheetViewSelect"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="sheetViewSelect"] [data-fc-value="current"]')
        ?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().sheetViews.activeViewId).toBeNull();

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="sheetViewSelect"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>(`[data-ribbon-command="sheetViewSelect"] [data-fc-value="${saved?.id}"]`)
        ?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().sheetViews.activeViewId).toBe(saved?.id);

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="sheetViewDelete"]')?.click();
      await flush();
    });

    state = mounted.instance.store.getState();
    expect(state.sheetViews.views).toEqual([]);
    expect(state.sheetViews.activeViewId).toBeNull();
  });

  it('creates a session chart from Insert > Chart menu instead of opening Quick Analysis', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const quickAnalysisSpy = vi.spyOn(mounted.instance, 'openQuickAnalysis');
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    const chartButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="chartInsert"] button',
    );
    expect(chartButton).toBeDefined();
    expect(mounted.instance.store.getState().charts.charts).toHaveLength(0);

    await act(async () => {
      chartButton?.click();
      await flush();
    });
    const column = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="column"]');
    await act(async () => {
      column?.click();
      await flush();
    });

    const charts = mounted.instance.store.getState().charts.charts;
    expect(charts).toHaveLength(1);
    expect(charts[0]).toMatchObject({
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      w: 360,
      h: 220,
    });
    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().charts.charts).toHaveLength(0);

    mounted.instance.history.redo();
    expect(mounted.instance.store.getState().charts.charts).toHaveLength(1);

    await act(async () => {
      chartButton?.click();
      await flush();
    });
    const scatter = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="scatter"]');
    expect(scatter?.textContent).toContain('Scatter');
    await act(async () => {
      scatter?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().charts.charts[1]).toMatchObject({
      kind: 'scatter',
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
    });
    expect(quickAnalysisSpy).not.toHaveBeenCalled();
  });

  it('creates a heuristic chart from Insert > Recommended Charts', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    const quickAnalysisSpy = vi.spyOn(mounted.instance, 'openQuickAnalysis');
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="chartInsert"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="recommended"]')?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().charts.charts[0]).toMatchObject({
      kind: 'pie',
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
    });
    expect(quickAnalysisSpy).not.toHaveBeenCalled();
  });

  it('routes Insert tab dialog commands to their workbook dialogs', async () => {
    mounted = await mountReactSpreadsheet();
    const pivot = vi.spyOn(mounted.instance, 'openPivotTableDialog');
    const names = vi.spyOn(mounted.instance, 'openNamedRangeDialog');
    const hyperlink = vi.spyOn(mounted.instance, 'openHyperlinkDialog');
    const links = vi.spyOn(mounted.instance, 'openExternalLinksDialog');
    const comment = vi.spyOn(mounted.instance, 'openCommentDialog');
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="pivotTableInsert"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="pivotTableInsert"] [data-cell-action="dialog"]')
        ?.click();
      await flush();
    });

    for (const command of ['linksInsert', 'commentInsert']) {
      await act(async () => {
        toolbar?.host
          .querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`)
          ?.click();
        await flush();
      });
    }
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="namedRangesInsert"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="namedRangesInsert"] [data-cell-action="define"]')
        ?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="hyperlinkInsert"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="hyperlinkInsert"] [data-cell-action="edit"]')
        ?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="hyperlinkInsert"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="hyperlinkInsert"] [data-cell-action="external"]')
        ?.click();
      await flush();
    });

    expect(pivot).toHaveBeenCalledTimes(1);
    expect(names).toHaveBeenCalledTimes(1);
    expect(hyperlink).toHaveBeenCalledTimes(1);
    expect(links).toHaveBeenCalledTimes(2);
    expect(comment).toHaveBeenCalledTimes(1);
  });

  it('opens Insert PivotTable menu recommended options with feedback', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="pivotTableInsert"] button')?.click();
      await flush();
    });

    const recommended = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotTableInsert"] [data-cell-action="recommended"]',
    );
    expect(recommended?.textContent).toContain('Recommended PivotTables');

    await act(async () => {
      recommended?.click();
      await flush();
    });

    const reportText = toolbar.host.querySelector('.demo__modal')?.textContent ?? '';
    expect(reportText).toContain('Recommended PivotTables');
    expect(reportText).toContain('PivotTable');
    expect(reportText).toContain('PivotCache and PivotTable mutation APIs');
  });

  it('opens Insert illustration menus with localized session feedback', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    const shapesButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="shapesInsert"] button',
    );
    await act(async () => {
      shapesButton?.click();
      await flush();
    });

    const roundedRectangle = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="shapesInsert"] [data-cell-action="rounded-rectangle"]',
    );
    expect(roundedRectangle?.textContent).toContain('Rounded Rectangle');

    await act(async () => {
      roundedRectangle?.click();
      await flush();
    });

    const reportText = toolbar.host.querySelector('.demo__modal')?.textContent ?? '';
    expect(reportText).toContain('Illustrations');
    expect(reportText).toContain('Rounded Rectangle');
    expect(reportText).toContain('OOXML chart, drawing, and media parts');
  });

  it('inserts a selected symbol from Insert > Symbol menu', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    const symbolButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"] button',
    );
    await act(async () => {
      symbolButton?.click();
      await flush();
    });
    const pi = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="π"]');
    await act(async () => {
      pi?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'π',
    });
  });

  it('inserts custom text from Insert > Symbol > More Symbols', async () => {
    mounted = await mountReactSpreadsheet();
    const originalPrompt = window.prompt;
    const prompt = vi.fn(() => 'Ωβ');
    Object.defineProperty(window, 'prompt', { configurable: true, value: prompt });
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    try {
      await act(async () => {
        toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="symbolInsert"] button')?.click();
        await flush();
      });
      const more = toolbar.host.querySelector<HTMLButtonElement>(
        '[data-cell-action="__more-symbols__"]',
      );
      await act(async () => {
        more?.click();
        await flush();
      });

      expect(prompt).toHaveBeenCalledWith('Character or Unicode text', '');
      expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'text',
        value: 'Ωβ',
      });
    } finally {
      Object.defineProperty(window, 'prompt', { configurable: true, value: originalPrompt });
    }
  });

  it('does not insert symbols into locked cells on protected sheets', async () => {
    mounted = await mountReactSpreadsheet();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setSheetProtected(mounted.instance.store, 0, true);
    toolbar = await renderToolbar(mounted, { activeTab: 'insert', onTabChange: vi.fn() });

    try {
      const symbolButton = toolbar.host.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="symbolInsert"] button',
      );
      await act(async () => {
        symbolButton?.click();
        await flush();
      });
      const yen = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="¥"]');
      await act(async () => {
        yen?.click();
        await flush();
      });

      expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(mounted.instance.history.canUndo()).toBe(false);
      expect(warn).toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it('undoes Review > Delete Comment as one format history action', async () => {
    mounted = await mountReactSpreadsheet();
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setActive(mounted.instance.store, addr);
    setComment(mounted.instance.store, addr, 'review note', mounted.instance.workbook);
    toolbar = await renderToolbar(mounted, { activeTab: 'review', onTabChange: vi.fn() });

    const deleteComment = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('button'),
    ).find((button) => button.getAttribute('aria-label') === 'Delete');
    expect(deleteComment).toBeDefined();

    await act(async () => {
      deleteComment?.click();
      await flush();
    });
    const deleteActive = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="delete-active"]',
    );
    await act(async () => {
      deleteActive?.click();
      await flush();
    });

    expect(commentAt(mounted.instance.store.getState(), addr)).toBeNull();
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(commentAt(mounted.instance.store.getState(), addr)).toBe('review note');

    mounted.instance.history.redo();
    expect(commentAt(mounted.instance.store.getState(), addr)).toBeNull();

    const a1 = { sheet: 0, row: 0, col: 0 };
    const c1 = { sheet: 0, row: 0, col: 2 };
    setComment(mounted.instance.store, a1, 'first', mounted.instance.workbook);
    setComment(mounted.instance.store, c1, 'second', mounted.instance.workbook);
    await act(async () => {
      deleteComment?.click();
      await flush();
    });
    const deleteAll = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="delete-all"]',
    );
    await act(async () => {
      deleteAll?.click();
      await flush();
    });
    expect(listComments(mounted.instance.store.getState())).toEqual([]);
    mounted.instance.history.undo();
    expect(listComments(mounted.instance.store.getState()).map((entry) => entry.text)).toEqual([
      'first',
      'second',
    ]);
  });

  it('does not create history when Review > Delete Comment targets an empty active cell', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 2, col: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'review', onTabChange: vi.fn() });

    const deleteComment = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('button'),
    ).find((button) => button.getAttribute('aria-label') === 'Delete');

    await act(async () => {
      deleteComment?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="delete-active"]')?.click();
      await flush();
    });

    expect(mounted.instance.history.canUndo()).toBe(false);
  });

  it('opens and navigates Review > Comments commands in row-major order', async () => {
    mounted = await mountReactSpreadsheet();
    const a1 = { sheet: 0, row: 0, col: 0 };
    const c1 = { sheet: 0, row: 0, col: 2 };
    const b3 = { sheet: 0, row: 2, col: 1 };
    setComment(mounted.instance.store, b3, 'third', mounted.instance.workbook);
    setComment(mounted.instance.store, a1, 'first', mounted.instance.workbook);
    setComment(mounted.instance.store, c1, 'second', mounted.instance.workbook);
    mutators.setActive(mounted.instance.store, a1);
    const openCommentDialog = vi.spyOn(mounted.instance, 'openCommentDialog');
    toolbar = await renderToolbar(mounted, { activeTab: 'review', onTabChange: vi.fn() });

    await act(async () => {
      toolbar?.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="newCommentReview"]')
        ?.click();
      await flush();
    });
    expect(openCommentDialog).toHaveBeenCalledTimes(1);

    await act(async () => {
      toolbar?.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="nextCommentReview"]')
        ?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().selection.active).toEqual(c1);

    await act(async () => {
      toolbar?.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="nextCommentReview"]')
        ?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().selection.active).toEqual(b3);

    await act(async () => {
      toolbar?.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="previousCommentReview"]')
        ?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().selection.active).toEqual(c1);
  });

  it('toggles sheet protection from Review and reflects the same state on View', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'review', onTabChange: vi.fn() });

    const reviewProtect = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectReview"]',
    );
    expect(reviewProtect?.getAttribute('aria-label')).toBe('Protect');
    const reviewWorkbookProtect = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectWorkbookReview"]',
    );
    expect(reviewWorkbookProtect?.getAttribute('aria-label')).toBe('Protect Workbook...');

    await act(async () => {
      reviewProtect?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);
    const reviewUnprotect = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectReview"]',
    );
    expect(reviewUnprotect?.className).toContain('demo__rb--active');
    expect(reviewUnprotect?.getAttribute('aria-label')).toBe('Unprotect');

    await act(async () => {
      reviewWorkbookProtect?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().protection.workbookStructure).toBeDefined();
    const reviewWorkbookUnprotect = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectWorkbookReview"]',
    );
    expect(reviewWorkbookUnprotect?.className).toContain('demo__rb--active');
    expect(reviewWorkbookUnprotect?.getAttribute('aria-label')).toBe('Unprotect Workbook...');

    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    const protectionMenu = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectionReview"] button',
    );
    await act(async () => {
      protectionMenu?.click();
      await flush();
    });
    const allowEditRange = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="allow-edit-range"]',
    );
    await act(async () => {
      allowEditRange?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().protection.allowedEditRanges).toMatchObject([
      { title: 'B2:C3', range: { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 } },
    ]);

    await act(async () => {
      protectionMenu?.click();
      await flush();
    });
    const clearAllowed = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="clear-allowed-edit-ranges"]',
    );
    await act(async () => {
      clearAllowed?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().protection.allowedEditRanges).toEqual([]);

    await toolbar.rerender('view');
    const viewProtect = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protect"]',
    );
    expect(viewProtect?.className).toContain('demo__rb--active');
    expect(viewProtect?.getAttribute('aria-label')).toBe('Unprotect');

    await act(async () => {
      viewProtect?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().protection.protectedSheets.has(0)).toBe(false);
    const viewUnprotected = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protect"]',
    );
    expect(viewUnprotected?.className).not.toContain('demo__rb--active');
    expect(viewUnprotected?.getAttribute('aria-label')).toBe('Protect');
  });

  it('marks the Cells > Format protect sheet item active when the sheet is protected', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.toggleSheetProtection();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const formatButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"] button',
    );
    await act(async () => {
      formatButton?.click();
      await flush();
    });

    const protectSheet = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="protectSheet"]',
    );
    expect(protectSheet?.className).toContain('demo__rb--active');
    expect(protectSheet?.getAttribute('role')).toBe('menuitemradio');
    expect(protectSheet?.getAttribute('aria-checked')).toBe('true');
  });

  it('routes Review > Find to the find tab of the Find and Replace dialog', async () => {
    mounted = await mountReactSpreadsheet();
    const openFindReplace = vi.spyOn(mounted.instance, 'openFindReplace');
    toolbar = await renderToolbar(mounted, { activeTab: 'review', onTabChange: vi.fn() });

    await act(async () => {
      toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="findReview"]')?.click();
      await flush();
    });

    expect(openFindReplace).toHaveBeenCalledWith('find');
  });

  it('runs built-in Review > Spelling when no host callback is provided', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'teh  report');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    toolbar = await renderToolbar(mounted, { activeTab: 'review', onTabChange: vi.fn() });

    const spelling = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.getAttribute('aria-label') === 'Spelling',
    );
    expect(spelling?.disabled).toBe(false);
    await act(async () => {
      spelling?.click();
      await flush();
    });

    expect(toolbar.host.textContent).toContain('Possible typo: "teh"');
  });

  it('runs built-in Review > Translate over the selected range when no host callback is provided', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'hello world');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'outside selection');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'review', onTabChange: vi.fn() });

    const translate = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.getAttribute('aria-label') === 'Translate',
    );
    expect(translate?.disabled).toBe(false);
    await act(async () => {
      translate?.click();
      await flush();
    });

    expect(toolbar.host.textContent).toContain('A1');
    expect(toolbar.host.textContent).toContain('Text ready for translation');
    expect(toolbar.host.textContent).not.toContain('outside selection');
  });

  it('runs built-in Automate > Script over the selected range', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'automate', onTabChange: vi.fn() });

    const script = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.getAttribute('aria-label') === 'Script',
    );
    expect(script?.disabled).toBe(false);
    await act(async () => {
      script?.click();
      await flush();
    });
    const commandSelect = toolbar.host.querySelector<HTMLSelectElement>('.demo__modal select');
    expect(commandSelect?.value).toBe('uppercase');
    const okButton = toolbar.host.querySelector<HTMLButtonElement>('.demo__btn--primary');
    await act(async () => {
      okButton?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'ALPHA',
    });
    expect(mounted.instance.history.canUndo()).toBe(true);
    const recordActions = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('button'),
    ).find((button) => button.getAttribute('aria-label') === 'Record Actions');
    await act(async () => {
      recordActions?.click();
      await flush();
    });
    expect(toolbar.host.textContent).toContain('Recorded selected range action');
    expect(toolbar.host.textContent).toContain('Uppercase ran on A1; 1 cell(s) changed.');
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('.demo__modal-x')?.click();
      await flush();
    });

    const allScripts = Array.from(toolbar.host.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.getAttribute('aria-label') === 'All Scripts',
    );
    await act(async () => {
      allScripts?.click();
      await flush();
    });
    expect(toolbar.host.textContent).toContain('Recent script runs');
    expect(toolbar.host.textContent).toContain('Script · 1 cell(s) changed');
    expect(toolbar.host.textContent).toContain('Uppercase ran on A1; 1 cell(s) changed.');

    mounted.instance.history.undo();
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
  });

  it('arms built-in border draw modes from Draw tab when no host callback is provided', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'draw', onTabChange: vi.fn() });

    const pen = toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="drawPen"]');
    const grid = toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="drawGrid"]');
    const eraser = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="drawErase"]',
    );
    expect(pen?.disabled).toBe(false);
    expect(grid?.disabled).toBe(false);
    expect(eraser?.disabled).toBe(false);

    await act(async () => {
      pen?.click();
      await flush();
    });
    expect(mounted.instance.borderDraw?.getMode()).toBe('draw');

    await act(async () => {
      grid?.click();
      await flush();
    });
    expect(mounted.instance.borderDraw?.getMode()).toBe('grid');

    await act(async () => {
      eraser?.click();
      await flush();
    });
    expect(mounted.instance.borderDraw?.getMode()).toBe('erase');
  });

  it('routes Formulas > Calculate Now to workbook recalculation', async () => {
    mounted = await mountReactSpreadsheet();
    const recalc = vi.spyOn(mounted.instance, 'recalc');
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    await act(async () => {
      toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="recalcNow"]')?.click();
      await flush();
    });

    expect(recalc).toHaveBeenCalledTimes(1);
  });

  it('sets and clears Page Layout > Print Area from the selection', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });

    const printAreaButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="printArea"] button',
    );
    await act(async () => {
      printAreaButton?.click();
      await flush();
    });
    const setPrintArea = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="set"]');
    await act(async () => {
      setPrintArea?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.printArea).toBe(
      'B2:C4',
    );

    await act(async () => {
      printAreaButton?.click();
      await flush();
    });
    const clearPrintArea = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="clear"]',
    );
    await act(async () => {
      clearPrintArea?.click();
      await flush();
    });
    expect(
      mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.printArea,
    ).toBeUndefined();
  });

  it('routes Page Layout > Print to the workbook print command', async () => {
    mounted = await mountReactSpreadsheet();
    const print = vi.spyOn(mounted.instance, 'print').mockImplementation(() => undefined);
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });

    const printPageLayout = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="printPageLayout"]',
    );
    expect(printPageLayout?.disabled).toBe(false);

    await act(async () => {
      printPageLayout?.click();
      await flush();
    });

    expect(print).toHaveBeenCalledTimes(1);
  });

  it('sets and clears Page Layout > Print Titles from the selection', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 3 });
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });

    const printTitlesButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="printTitles"] button',
    );
    await act(async () => {
      printTitlesButton?.click();
      await flush();
    });
    const rows = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="rows"]');
    await act(async () => {
      rows?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.printTitleRows).toBe(
      '1:2',
    );

    await act(async () => {
      printTitlesButton?.click();
      await flush();
    });
    const cols = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="cols"]');
    await act(async () => {
      cols?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.printTitleCols).toBe(
      'B:D',
    );

    await act(async () => {
      printTitlesButton?.click();
      await flush();
    });
    const clear = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="clear"]');
    await act(async () => {
      clear?.click();
      await flush();
    });
    const setup = mounted.instance.store.getState().pageSetup.setupBySheet.get(0);
    expect(setup?.printTitleRows).toBeUndefined();
    expect(setup?.printTitleCols).toBeUndefined();
  });

  it('sets and clears Page Layout > Breaks from the active cell', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 4, col: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });

    const breaksButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageBreaks"] button',
    );
    await act(async () => {
      breaksButton?.click();
      await flush();
    });
    const insertRow = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="insert-row"]',
    );
    await act(async () => {
      insertRow?.click();
      await flush();
    });
    expect(
      mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.manualPageBreakRows,
    ).toEqual([4]);

    await act(async () => {
      breaksButton?.click();
      await flush();
    });
    const insertCol = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="insert-col"]',
    );
    await act(async () => {
      insertCol?.click();
      await flush();
    });
    expect(
      mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.manualPageBreakCols,
    ).toEqual([2]);

    await act(async () => {
      breaksButton?.click();
      await flush();
    });
    const reset = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="reset"]');
    await act(async () => {
      reset?.click();
      await flush();
    });
    const setup = mounted.instance.store.getState().pageSetup.setupBySheet.get(0);
    expect(setup?.manualPageBreakRows).toBeUndefined();
    expect(setup?.manualPageBreakCols).toBeUndefined();
  });

  it('sets and clears Page Layout > Background from an image file', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });
    class TestFileReader {
      result: string | null = null;
      onload: ((this: FileReader, event: ProgressEvent<FileReader>) => void) | null = null;

      readAsDataURL(): void {
        this.result = 'data:image/png;base64,YWJj';
        this.onload?.call(
          this as unknown as FileReader,
          new ProgressEvent('load') as ProgressEvent<FileReader>,
        );
      }
    }
    vi.stubGlobal('FileReader', TestFileReader);

    const backgroundButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sheetBackground"] button',
    );
    await act(async () => {
      backgroundButton?.click();
      await flush();
    });
    const setBackground = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="set"]');
    await act(async () => {
      setBackground?.click();
      await flush();
    });
    const input = toolbar.host.querySelector<HTMLInputElement>(
      '[data-ribbon-file-input="sheetBackground"]',
    );
    expect(input).not.toBeNull();
    Object.defineProperty(input, 'files', {
      configurable: true,
      value: [new File(['abc'], 'bg.png', { type: 'image/png' })],
    });
    await act(async () => {
      input?.dispatchEvent(new Event('change', { bubbles: true }));
      await flush();
      await new Promise((resolve) => setTimeout(resolve, 0));
      await flush();
    });
    expect(mounted.instance.store.getState().ui.sheetBackgroundImages.get(0)).toBe(
      'data:image/png;base64,YWJj',
    );
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);
    expect(mounted.instance.history.redo()).toBe(true);
    expect(mounted.instance.store.getState().ui.sheetBackgroundImages.get(0)).toBe(
      'data:image/png;base64,YWJj',
    );

    await act(async () => {
      backgroundButton?.click();
      await flush();
    });
    const clearBackground = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="clear"]',
    );
    await act(async () => {
      clearBackground?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().ui.sheetBackgroundImages.get(0)).toBe(
      'data:image/png;base64,YWJj',
    );
  });

  it('toggles Formula Auditing > Show Formulas from the Formulas tab', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    const showFormulas = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="showFormulasFormula"]',
    );
    expect(showFormulas).not.toBeNull();
    expect(showFormulas?.className).not.toContain('demo__rb--active');

    await act(async () => {
      showFormulas?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().ui.showFormulas).toBe(true);
    expect(showFormulas?.className).toContain('demo__rb--active');
  });

  it('runs Formulas > Formula Auditing > Error Checking against formula errors', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 1 },
      { kind: 'error', code: 7, text: '#DIV/0!' },
      '=1/0',
    );
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    const errorCheckingButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="errorChecking"] button',
    );
    await act(async () => {
      errorCheckingButton?.click();
      await flush();
    });
    const errorChecking = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="errorChecking"]',
    );
    await act(async () => {
      errorChecking?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 0,
      col: 1,
    });

    await act(async () => {
      errorCheckingButton?.click();
      await flush();
    });
    const ignoreError = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="ignoreError"]',
    );
    await act(async () => {
      ignoreError?.click();
      await flush();
    });

    expect([...mounted.instance.store.getState().errorIndicators.ignoredErrors]).toEqual([
      '0:0:1',
    ]);

    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().errorIndicators.ignoredErrors.size).toBe(0);
    expect(mounted.instance.history.redo()).toBe(true);
    expect([...mounted.instance.store.getState().errorIndicators.ignoredErrors]).toEqual([
      '0:0:1',
    ]);
  });

  it('runs Formula Auditing trace arrows from the Formulas tab with undoable state', async () => {
    mounted = await mountReactSpreadsheet();
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b1 = { sheet: 0, row: 0, col: 1 };
    mounted.instance.workbook.setNumber(a1, 10);
    mounted.instance.workbook.setFormula(b1, '=A1');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setActive(mounted.instance.store, b1);
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="precedents"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().traces.items).toEqual([
      { kind: 'precedent', from: a1, to: b1 },
    ]);

    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().traces.items).toEqual([]);
    expect(mounted.instance.history.redo()).toBe(true);
    expect(mounted.instance.store.getState().traces.items).toEqual([
      { kind: 'precedent', from: a1, to: b1 },
    ]);

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="clearArrows"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="clear-all"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().traces.items).toEqual([]);
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().traces.items).toEqual([
      { kind: 'precedent', from: a1, to: b1 },
    ]);

    mounted.instance.clearTraces();
    mutators.setActive(mounted.instance.store, a1);
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="dependents"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().traces.items).toEqual([
      { kind: 'dependent', from: a1, to: b1 },
    ]);

    mutators.setActive(mounted.instance.store, b1);
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="precedents"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().traces.items).toEqual([
      { kind: 'dependent', from: a1, to: b1 },
      { kind: 'precedent', from: a1, to: b1 },
    ]);
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="clearArrows"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="clear-precedents"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().traces.items).toEqual([
      { kind: 'dependent', from: a1, to: b1 },
    ]);
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="clearArrows"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="clear-dependents"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().traces.items).toEqual([]);
  });

  it('opens Formulas > Formula Auditing > Evaluate Formula dialog', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'number', value: 2 },
      '=1+1',
    );
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    const evaluateFormula = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="evaluateFormula"]',
    );
    await act(async () => {
      evaluateFormula?.click();
      await flush();
    });

    expect(document.querySelector<HTMLElement>('.fc-evaldlg')?.hidden).toBe(false);
    expect(document.querySelector<HTMLElement>('.fc-evaldlg__box')?.textContent).toBe('=1+1');
  });

  it('opens the default-off Watch Window from the Formulas tab on demand', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'number', value: 2 },
      '=1+1',
    );
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    expect(mounted.instance.features.watchWindow).toBeUndefined();
    const watch = toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="watch"] button');
    await act(async () => {
      watch?.click();
      await flush();
    });
    const openWatch = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="open"]');
    await act(async () => {
      openWatch?.click();
      await flush();
    });

    expect(mounted.instance.features.watchWindow).toBeDefined();
    expect(mounted.instance.store.getState().ui.watchPanelOpen).toBe(true);
    await act(async () => {
      watch?.click();
      await flush();
    });
    const addWatch = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="add"]');
    await act(async () => {
      addWatch?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().watch.watches).toEqual([{ sheet: 0, row: 0, col: 0 }]);
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().watch.watches).toEqual([]);
    expect(mounted.instance.history.redo()).toBe(true);
    expect(mounted.instance.store.getState().watch.watches).toEqual([{ sheet: 0, row: 0, col: 0 }]);
    await act(async () => {
      watch?.click();
      await flush();
    });
    const deleteWatch = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="delete"]');
    await act(async () => {
      deleteWatch?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().watch.watches).toEqual([]);
    await act(async () => {
      watch?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="add"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().watch.watches).toEqual([{ sheet: 0, row: 0, col: 0 }]);
    await act(async () => {
      watch?.click();
      await flush();
    });
    const deleteAll = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="delete-all"]');
    await act(async () => {
      deleteAll?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().watch.watches).toEqual([]);
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.store.getState().watch.watches).toEqual([{ sheet: 0, row: 0, col: 0 }]);
  });

  it('sets Page Layout scale-to-fit values directly from the ribbon', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });

    const widthButton = toolbar.host.querySelector<HTMLButtonElement>('button[aria-label="Width"]');
    await act(async () => {
      widthButton?.click();
      await flush();
    });
    const widthList = toolbar.host.querySelector<HTMLElement>(
      '[role="listbox"][aria-label="Width"]',
    );
    const onePage = Array.from(
      widthList?.querySelectorAll<HTMLButtonElement>('[role="option"]') ?? [],
    ).find((button) => button.dataset.fcValue === '1');
    await act(async () => {
      onePage?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.fitWidth).toBe(1);

    const scaleButton = toolbar.host.querySelector<HTMLButtonElement>('button[aria-label="Scale"]');
    expect(scaleButton).not.toBeNull();
    expect(scaleButton?.disabled).toBe(false);
    await act(async () => {
      scaleButton?.click();
      await flush();
    });
    const scaleList = toolbar.host.querySelector<HTMLElement>(
      '[role="listbox"][aria-label="Scale"]',
    );
    expect(scaleList).not.toBeNull();
    const seventyFive = Array.from(
      scaleList?.querySelectorAll<HTMLButtonElement>('[role="option"]') ?? [],
    ).find((button) => button.dataset.fcValue === '75');
    await act(async () => {
      seventyFive?.click();
      await flush();
    });

    const setup = mounted.instance.store.getState().pageSetup.setupBySheet.get(0);
    expect(setup?.fitWidth).toBeUndefined();
    expect(setup?.fitHeight).toBeUndefined();
    expect(setup?.scale).toBe(0.75);
  });

  it('switches the workbook host theme from Page Layout > Themes', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });

    const themeButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageTheme"] button',
    );
    await act(async () => {
      themeButton?.click();
      await flush();
    });
    const ink = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="ink"]');
    await act(async () => {
      ink?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().ui.theme).toBe('ink');
    await act(async () => {
      themeButton?.click();
      await flush();
    });
    expect(
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-cell-action="ink"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
  });

  it('toggles Page Layout sheet options for view and print', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'pageLayout', onTabChange: vi.fn() });

    const gridView = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageLayoutGridlinesView"]',
    );
    const gridPrint = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageLayoutGridlinesPrint"]',
    );
    const headingView = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageLayoutHeadingsView"]',
    );
    const headingPrint = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageLayoutHeadingsPrint"]',
    );

    await act(async () => {
      gridView?.click();
      gridPrint?.click();
      headingView?.click();
      headingPrint?.click();
      await flush();
    });

    const state = mounted.instance.store.getState();
    expect(state.ui.showGridLines).toBe(false);
    expect(state.ui.showHeaders).toBe(false);
    const setup = state.pageSetup.setupBySheet.get(0);
    expect(setup?.showGridlines).toBe(true);
    expect(setup?.showHeadings).toBe(true);

    await act(async () => {
      expect(mounted.instance.history.undo()).toBe(true);
      await flush();
    });
    expect(mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.showHeadings).not.toBe(
      true,
    );
    await act(async () => {
      expect(mounted.instance.history.undo()).toBe(true);
      await flush();
    });
    expect(
      mounted.instance.store.getState().pageSetup.setupBySheet.get(0)?.showGridlines,
    ).not.toBe(true);
  });

  it('routes Data > Data Validation through the dedicated instance API', async () => {
    mounted = await mountReactSpreadsheet();
    const dataValidationSpy = vi.spyOn(mounted.instance, 'openDataValidationDialog');
    const formatDialogSpy = vi.spyOn(mounted.instance, 'openFormatDialog');
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const dataValidationButton = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('button'),
    ).find((b) => b.getAttribute('aria-label') === 'Data Validation');
    expect(dataValidationButton).toBeDefined();

    await act(async () => {
      dataValidationButton?.click();
      await flush();
    });
    const settings = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="settings"]');
    await act(async () => {
      settings?.click();
      await flush();
    });

    expect(dataValidationSpy).toHaveBeenCalledTimes(1);
    expect(formatDialogSpy).not.toHaveBeenCalled();
  });

  it('clears validation from the selected cells from Data > Data Validation menu', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { bold: true, validation: { kind: 'list', source: ['A', 'B'] } },
    );
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const dataValidationButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="dataValidation"] button',
    );
    await act(async () => {
      dataValidationButton?.click();
      await flush();
    });
    const clearValidation = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="clearValidation"]',
    );
    await act(async () => {
      clearValidation?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().format.formats.get('0:0:0')).toEqual({ bold: true });
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().format.formats.get('0:0:0')).toMatchObject({
      bold: true,
      validation: { kind: 'list', source: ['A', 'B'] },
    });

    mounted.instance.history.redo();
    expect(mounted.instance.store.getState().format.formats.get('0:0:0')).toEqual({ bold: true });
  });

  it('circles and clears invalid validation cells from Data > Data Validation menu', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 1 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 0, row: 8, col: 3 },
      { validation: { kind: 'list', source: ['Open', 'Closed'] } },
    );
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 1, row: 0, col: 0 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'number', value: 99 },
    );
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 1 },
      { kind: 'number', value: 5 },
    );
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 8, col: 3 },
      { kind: 'text', value: 'Hold' },
    );
    mutators.setCell(
      mounted.instance.store,
      { sheet: 1, row: 0, col: 0 },
      { kind: 'number', value: 99 },
    );
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const dataValidationButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="dataValidation"] button',
    );
    await act(async () => {
      dataValidationButton?.click();
      await flush();
    });
    const circleInvalid = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="circleInvalid"]',
    );
    await act(async () => {
      circleInvalid?.click();
      await flush();
    });

    expect([...mounted.instance.store.getState().errorIndicators.validationCircles]).toEqual([
      '0:0:0',
      '0:8:3',
    ]);
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().errorIndicators.validationCircles.size).toBe(0);
    mounted.instance.history.redo();
    expect([...mounted.instance.store.getState().errorIndicators.validationCircles]).toEqual([
      '0:0:0',
      '0:8:3',
    ]);

    await act(async () => {
      dataValidationButton?.click();
      await flush();
    });
    const clearCircles = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="clearCircles"]',
    );
    await act(async () => {
      clearCircles?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().errorIndicators.validationCircles.size).toBe(0);

    mounted.instance.history.undo();
    expect([...mounted.instance.store.getState().errorIndicators.validationCircles]).toEqual([
      '0:0:0',
      '0:8:3',
    ]);
  });

  it('opens the conditional-formatting menu and applies preset rules', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const conditionalButton = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('button'),
    ).find((b) => b.getAttribute('aria-label') === 'Conditional formatting');
    expect(conditionalButton).toBeDefined();

    await act(async () => {
      conditionalButton?.click();
      await flush();
    });

    expect(toolbar.host.querySelector('.demo__cf-menu')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="cell-greater"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="cell-less"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="cell-between"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="cell-equal"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="text-contains"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="date-occurring"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="unique"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="top10-percent"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="data-solid-gray"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="scale-gwg"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="icons-arrows5"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="highlight-more"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="top-bottom-more"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="data-bars-more"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="color-scales-more"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="icon-sets-more"]')).toBeTruthy();
    expect(
      Array.from(toolbar.host.querySelectorAll('.demo__cf-menu__submenu')).some((el) =>
        el.textContent?.includes('Clear Rules'),
      ),
    ).toBe(true);
    expect(toolbar.host.querySelector('[data-cf-action="clear-selection"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cf-action="clear-sheet"]')).toBeTruthy();
    expect(
      toolbar.host.querySelector('[data-cf-action="data-solid-green"]')?.getAttribute('title'),
    ).toBe('Solid Fill, Green Data Bar');
    expect(
      toolbar.host.querySelector('[data-cf-action="scale-gwg"]')?.getAttribute('aria-label'),
    ).toBe('Green - White - Green Color Scale');
    expect(
      toolbar.host.querySelector('[data-cf-action="icons-arrows5"]')?.getAttribute('aria-label'),
    ).toBe('5 Arrows');
    expect(toolbar.host.querySelector('[data-cf-action="icons-arrows5"]')?.className).toContain(
      'demo__cf-menu__iconset',
    );
    const preset = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cf-action="data-solid-green"]',
    );
    expect(preset).toBeDefined();

    await act(async () => {
      preset?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().conditional.rules).toEqual([
      {
        kind: 'data-bar',
        range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 },
        color: '#70ad47',
        gradient: false,
        showValue: true,
      },
    ]);
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().conditional.rules).toEqual([]);

    mounted.instance.history.redo();
    expect(mounted.instance.store.getState().conditional.rules).toHaveLength(1);
  });

  it('opens conditional-formatting Other Rules entries with the matching rule kind', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });
    const openConditionalDialog = vi
      .spyOn(mounted.instance, 'openConditionalDialog')
      .mockImplementation(() => undefined);

    const conditionalButton = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('button'),
    ).find((b) => b.getAttribute('aria-label') === 'Conditional formatting');

    await act(async () => {
      conditionalButton?.click();
      await flush();
    });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cf-action="data-bars-more"]')?.click();
      await flush();
    });

    expect(openConditionalDialog).toHaveBeenCalledWith({ mode: 'new', kind: 'data-bar' });

    await act(async () => {
      conditionalButton?.click();
      await flush();
    });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cf-action="icon-sets-more"]')?.click();
      await flush();
    });

    expect(openConditionalDialog).toHaveBeenLastCalledWith({ mode: 'new', kind: 'icon-set' });
  });

  it('marks the conditional-formatting button active when the selection has rules', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    mutators.addConditionalRule(mounted.instance.store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 3 },
      color: '#70ad47',
      showValue: true,
    });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const conditionalButton = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('button'),
    ).find((b) => b.getAttribute('aria-label') === 'Conditional formatting');

    expect(conditionalButton?.className).toContain('demo__rb--active');
  });

  it('applies a cell style from the Home > Cell Styles menu', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const cellStylesButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"] button',
    );
    await act(async () => {
      cellStylesButton?.click();
      await flush();
    });

    const good = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="good"]');
    await act(async () => {
      good?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().format.formats.get('0:0:0')).toMatchObject({
      color: '#006100',
      fill: '#c6efce',
      cellStyle: 'good',
    });
  });

  it('marks the active Cell Styles item from the active cell style id', async () => {
    mounted = await mountReactSpreadsheet();
    applyCellStyle(mounted.instance.store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'good');
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const cellStylesButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="cellStyles"] button',
    );
    expect(cellStylesButton?.className).toContain('demo__rb--active');

    await act(async () => {
      cellStylesButton?.click();
      await flush();
    });

    expect(toolbar.host.textContent).toContain('Good, Bad and Neutral');
    expect(toolbar.host.textContent).toContain('Data and Model');
    const good = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="good"]');
    expect(good?.getAttribute('role')).toBe('menuitemradio');
    expect(good?.getAttribute('aria-checked')).toBe('true');
    expect(good?.className).toContain('demo__rb--active');
  });

  it('applies a Format as Table style from the Home menu', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const formatTableButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableHome"] button',
    );
    await act(async () => {
      formatTableButton?.click();
      await flush();
    });

    const dark = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="dark"]');
    await act(async () => {
      dark?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().tables.tables).toEqual([
      {
        id: 'table-0-0-0-3-2',
        source: 'session',
        range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
        style: 'dark',
        color: '#5b9bd5',
        showHeader: true,
        showTotal: false,
        banded: true,
      },
    ]);
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().tables.tables).toEqual([]);

    mounted.instance.history.redo();
    expect(mounted.instance.store.getState().tables.tables).toHaveLength(1);
  });

  it('marks Format as Table active when the active cell is inside a table', async () => {
    mounted = await mountReactSpreadsheet();
    formatAsTable(
      mounted.instance.store,
      { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 },
      { style: 'medium' },
    );
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 2, col: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const formatTableButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableHome"] button',
    );
    expect(formatTableButton?.className).toContain('demo__rb--active');
  });

  it('opens Cells menus and runs cell-shift / hide actions', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 1 }, 'B2');
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const insertButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"] button',
    );
    expect(insertButton).toBeDefined();
    await act(async () => {
      insertButton?.click();
      await flush();
    });
    const shiftDown = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="shiftDown"]',
    );
    await act(async () => {
      shiftDown?.click();
      await flush();
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'text',
      value: 'B2',
    });
    mounted.instance.history.undo();
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'B2',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 }).kind).toBe('blank');

    mounted.instance.history.redo();
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'text',
      value: 'B2',
    });

    const formatButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"] button',
    );
    await act(async () => {
      formatButton?.click();
      await flush();
    });
    const hideRows = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="hideRows"]');
    await act(async () => {
      hideRows?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().layout.hiddenRows.has(1)).toBe(false);

    mounted.instance.history.redo();
    expect(mounted.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);
  });

  it('undoes Delete Cells menu shift-left as one ribbon action', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 1 }, 'B2');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 2 }, 'C2');
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const deleteButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"] button',
    );
    expect(deleteButton).toBeDefined();
    await act(async () => {
      deleteButton?.click();
      await flush();
    });

    const shiftLeft = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="shiftLeft"]',
    );
    await act(async () => {
      shiftLeft?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'C2',
    });

    mounted.instance.history.undo();
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'B2',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({
      kind: 'text',
      value: 'C2',
    });

    mounted.instance.history.redo();
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'C2',
    });
  });

  it('runs a custom Sort & Filter sort from the sort dialog', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'paper');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 3);
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'ink');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'eraser');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const sortButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"] button',
    );
    await act(async () => {
      sortButton?.click();
      await flush();
    });
    const customSort = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="custom"]');
    await act(async () => {
      customSort?.click();
      await flush();
    });
    const [columnSelect] = Array.from(
      toolbar.host.querySelectorAll<HTMLSelectElement>('.demo__sort-dialog select'),
    );
    const headerCheckbox = toolbar.host.querySelector<HTMLInputElement>(
      '.demo__sort-dialog__check input',
    );
    await act(async () => {
      if (columnSelect) {
        columnSelect.value = '1';
        columnSelect.dispatchEvent(new Event('change', { bubbles: true }));
      }
      headerCheckbox?.click();
      await flush();
    });
    const okButton = Array.from(
      toolbar.host.querySelectorAll<HTMLButtonElement>('.demo__btn'),
    ).find((button) => button.textContent === 'OK');
    await act(async () => {
      okButton?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'ink',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'eraser',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'text',
      value: 'paper',
    });
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'paper',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'ink',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'text',
      value: 'eraser',
    });

    mounted.instance.history.redo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'ink',
    });
  });

  it('sorts the surrounding current region from a single active cell', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'item');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'qty');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'paper');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 3);
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'ink');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 1 }, 1);
    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 0 }, 'eraser');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 3, col: 1 }, 2);
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 2, col: 1 });
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const sortButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"] button',
    );
    await act(async () => {
      sortButton?.click();
      await flush();
    });
    const ascending = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="asc"]');
    await act(async () => {
      ascending?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'item',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'ink',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'text',
      value: 'eraser',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({
      kind: 'text',
      value: 'paper',
    });
  });

  it('runs Data > Sort Descending against the inferred current region', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'low');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 3);
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 1 }, 'high');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 0 }, 2);
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 1 }, 'mid');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const sortDesc = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortDesc"]',
    );
    expect(sortDesc?.disabled).toBe(false);

    await act(async () => {
      sortDesc?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 3,
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'high',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'text',
      value: 'low',
    });

    mounted.instance.history.undo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
  });

  it('opens the custom sort dialog from Data > Sort', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'paper');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 3);
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'ink');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const sortButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortData"] button',
    );
    expect(sortButton?.disabled).toBe(false);

    await act(async () => {
      sortButton?.click();
      await flush();
    });
    const customSort = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="custom"]');
    await act(async () => {
      customSort?.click();
      await flush();
    });

    expect(toolbar.host.querySelector('.demo__sort-dialog')).toBeTruthy();
  });

  it("filters by the selected cell's value from the Sort & Filter menu", async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'item');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'paper');
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'ink');
    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 0 }, 'paper');
    mounted.instance.workbook.setText({ sheet: 0, row: 4, col: 0 }, 'ink');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 2, col: 0 });
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const sortButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"] button',
    );
    await act(async () => {
      sortButton?.click();
      await flush();
    });
    const filterBySelected = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="filter-by-selected"]',
    );
    await act(async () => {
      filterBySelected?.click();
      await flush();
    });

    const s = mounted.instance.store.getState();
    expect(s.layout.hiddenRows.has(1)).toBe(true);
    expect(s.layout.hiddenRows.has(2)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(true);
    expect(s.layout.hiddenRows.has(4)).toBe(false);
    expect(s.ui.filterCriteria).toEqual([
      {
        range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
        byCol: 0,
        hiddenValues: ['paper'],
      },
    ]);

    await act(async () => {
      sortButton?.click();
      await flush();
    });
    expect(toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="filter-reapply"]')).not.toBeNull();
    const advanced = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="filter-advanced"]',
    );
    await act(async () => {
      advanced?.click();
      await flush();
    });
    expect(toolbar.host.querySelector('.demo__modal')?.textContent).toContain('Advanced Filter');
  });

  it('sets row height and column width from the Cells > Format menu', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 2, r1: 2, c1: 3 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const formatButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"] button',
    );
    await act(async () => {
      formatButton?.click();
      await flush();
    });
    const rowHeight = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="rowHeight"]',
    );
    await act(async () => {
      rowHeight?.click();
      await flush();
    });
    const rowHeightInput = toolbar.host.querySelector<HTMLInputElement>(
      '.demo__modal input[type="number"]',
    );
    await act(async () => {
      if (rowHeightInput) {
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
          rowHeightInput,
          '33',
        );
        rowHeightInput.dispatchEvent(new Event('input', { bubbles: true }));
      }
      toolbar.host.querySelector<HTMLButtonElement>('.demo__btn--primary')?.click();
      await flush();
    });

    await act(async () => {
      formatButton?.click();
      await flush();
    });
    const colWidth = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="colWidth"]');
    await act(async () => {
      colWidth?.click();
      await flush();
    });
    const colWidthInput = toolbar.host.querySelector<HTMLInputElement>(
      '.demo__modal input[type="number"]',
    );
    await act(async () => {
      if (colWidthInput) {
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
          colWidthInput,
          '144',
        );
        colWidthInput.dispatchEvent(new Event('input', { bubbles: true }));
      }
      toolbar.host.querySelector<HTMLButtonElement>('.demo__btn--primary')?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().layout.rowHeights.get(1)).toBe(33);
    expect(mounted.instance.store.getState().layout.rowHeights.get(2)).toBe(33);
    expect(mounted.instance.store.getState().layout.colWidths.get(2)).toBe(144);
    expect(mounted.instance.store.getState().layout.colWidths.get(3)).toBe(144);
  });

  it('runs sheet rename and hide commands from the Cells > Format menu', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.addSheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const openFormatMenu = async (): Promise<void> => {
      const formatButton = toolbar.host.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="formatCellsHome"] button',
      );
      await act(async () => {
        formatButton?.click();
        await flush();
      });
    };
    await openFormatMenu();
    expect(toolbar.host.querySelector('[data-cell-action="renameSheet"]')).toBeTruthy();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="hideSheet"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.hiddenSheets.has(0)).toBe(true);
    expect(mounted.instance.store.getState().data.sheetIndex).toBe(1);

    await openFormatMenu();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="unhideSheet"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.hiddenSheets.has(0)).toBe(false);

    await openFormatMenu();
    expect(toolbar.host.querySelector('[data-cell-action="moveSheetLeft"]')).toBeTruthy();
    expect(toolbar.host.querySelector('[data-cell-action="moveSheetRight"]')).toBeTruthy();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="moveSheetLeft"]')?.click();
      await flush();
    });
    if (mounted.instance.workbook.capabilities.sheetMutate) {
      expect(mounted.instance.store.getState().data.sheetIndex).toBe(0);
    }
    await openFormatMenu();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="moveSheetRight"]')?.click();
      await flush();
    });
    if (mounted.instance.workbook.capabilities.sheetMutate) {
      expect(mounted.instance.store.getState().data.sheetIndex).toBe(1);
    }

    await openFormatMenu();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="tabColorBlue"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.sheetTabColors.get(1)).toBe('#4472c4');
    await openFormatMenu();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="tabColorNone"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.sheetTabColors.has(1)).toBe(false);

    await openFormatMenu();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="renameSheet"]')?.click();
      await flush();
    });
    expect(toolbar.host.querySelector('.demo__modal')?.textContent).toContain('Rename');
    const renameInput = toolbar.host.querySelector<HTMLInputElement>('.demo__modal input');
    await act(async () => {
      if (renameInput) {
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
          renameInput,
          'Renamed',
        );
        renameInput.dispatchEvent(new Event('input', { bubbles: true }));
        renameInput.dispatchEvent(new Event('change', { bubbles: true }));
      }
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('.demo__btn--primary')?.click();
      await flush();
    });
    expect(toolbar.host.querySelector('.demo__modal')).toBeNull();
    if (mounted.instance.workbook.capabilities.sheetMutate) {
      expect(mounted.instance.workbook.sheetName(1)).toBe('Renamed');
    }
  });

  it('inserts and deletes sheets from the Cells insert/delete menus', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const insertButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"] button',
    );
    await act(async () => {
      insertButton?.click();
      await flush();
    });
    expect(toolbar.host.querySelector('[data-cell-action="sheet"]')).toBeTruthy();
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="sheet"]')?.click();
      await flush();
    });
    if (mounted.instance.workbook.capabilities.sheetMutate) {
      expect(mounted.instance.workbook.sheetCount).toBe(2);
      expect(mounted.instance.store.getState().data.sheetIndex).toBe(1);
    }

    const deleteButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"] button',
    );
    await act(async () => {
      deleteButton?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="sheet"]')?.click();
      await flush();
    });
    if (mounted.instance.workbook.capabilities.sheetMutate) {
      expect(mounted.instance.workbook.sheetCount).toBe(1);
      expect(mounted.instance.store.getState().data.sheetIndex).toBe(0);
    }
  });

  it('autofits row height and column width from the Cells > Format menu', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 2 }, 'wide enough for autofit');
    mounted.instance.workbook.setText(
      { sheet: 0, row: 2, col: 2 },
      'first line\nsecond line\nthird line',
    );
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 2, r1: 2, c1: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const formatButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"] button',
    );
    await act(async () => {
      formatButton?.click();
      await flush();
    });
    toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="autoFitColWidth"]')?.click();

    await act(async () => {
      formatButton?.click();
      await flush();
    });
    toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="autoFitRowHeight"]')?.click();

    expect(mounted.instance.store.getState().layout.colWidths.get(2)).toBeGreaterThan(100);
    expect(mounted.instance.store.getState().layout.rowHeights.get(2)).toBeGreaterThan(40);
    expect(mounted.instance.history.canUndo()).toBe(true);
  });

  it('opens Editing menus and runs fill / clear actions', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'A1');
    mounted.instance.history.clear();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });
    expect(toolbar.host.querySelectorAll('[data-ribbon-command="clearFormat"]')).toHaveLength(1);

    const fillButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="fillHome"] button',
    );
    await act(async () => {
      fillButton?.click();
      await flush();
    });
    const fillDown = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="down"]');
    await act(async () => {
      fillDown?.click();
      await flush();
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'A2',
    });
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
    expect(mounted.instance.history.redo()).toBe(true);
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'A2',
    });

    const clearButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="clearFormat"] button',
    );
    await act(async () => {
      clearButton?.click();
      await flush();
    });
    const clearContents = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="contents"]',
    );
    await act(async () => {
      clearContents?.click();
      await flush();
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'A2',
    });
  });

  it('runs Flash Fill from the Fill menu using adjacent-column examples', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'John Smith');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'John');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'Jane Doe');
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'Bob Lee');
    mounted.instance.history.clear();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="flash"]')?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'Jane',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'text',
      value: 'Bob',
    });
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 }).kind).toBe('blank');
  });

  it('skips locked cells when Flash Fill runs on a protected sheet', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'John Smith');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'John');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'Jane Doe');
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'Bob Lee');
    mutators.setCellFormat(mounted.instance.store, { sheet: 0, row: 2, col: 1 }, { locked: false });
    mutators.setSheetProtected(mounted.instance.store, 0, true);
    mounted.instance.history.clear();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 1, r1: 2, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="flash"]')?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'text',
      value: 'Bob',
    });
    expect(mounted.instance.history.canUndo()).toBe(true);
  });

  it('routes Sort & Filter conditional formatting to the rules manager', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });
    const openCfRulesDialog = vi
      .spyOn(mounted.instance, 'openCfRulesDialog')
      .mockImplementation(() => undefined);
    const openConditionalDialog = vi
      .spyOn(mounted.instance, 'openConditionalDialog')
      .mockImplementation(() => undefined);

    const sortButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"] button',
    );
    await act(async () => {
      sortButton?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="conditional"]')?.click();
      await flush();
    });

    expect(openCfRulesDialog).toHaveBeenCalledTimes(1);
    expect(openConditionalDialog).not.toHaveBeenCalled();
  });

  it('clears contents only in unlocked cells on protected sheets', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'locked');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'unlocked');
    mutators.setCellFormat(mounted.instance.store, { sheet: 0, row: 0, col: 1 }, { locked: false });
    mutators.setSheetProtected(mounted.instance.store, 0, true);
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const clearButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="clearFormat"] button',
    );
    await act(async () => {
      clearButton?.click();
      await flush();
    });
    const clearContents = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="contents"]',
    );
    await act(async () => {
      clearContents?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'locked',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'unlocked',
    });
  });

  it('clears visual formats without removing comments, hyperlinks, or validation', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        fill: '#ff0000',
        comment: 'keep',
        hyperlink: 'https://example.com',
        validation: { kind: 'list', source: ['A', 'B'] },
      },
    );
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const clearButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="clearFormat"] button',
    );
    await act(async () => {
      clearButton?.click();
      await flush();
    });
    const clearFormats = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="formats"]',
    );
    await act(async () => {
      clearFormats?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().format.formats.get('0:0:0')).toEqual({
      comment: 'keep',
      hyperlink: 'https://example.com',
      validation: { kind: 'list', source: ['A', 'B'] },
    });
  });

  it('clears conditional formats from Home > Clear with undo support', async () => {
    mounted = await mountReactSpreadsheet();
    const rule: ConditionalRule = {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      op: '>',
      a: 10,
      apply: { fill: '#fff2cc' },
    };
    mutators.addConditionalRule(mounted.instance.store, rule);
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const clearButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="clearFormat"] button',
    );
    await act(async () => {
      clearButton?.click();
      await flush();
    });
    const clearConditional = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="conditional"]',
    );
    await act(async () => {
      clearConditional?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().conditional.rules).toEqual([]);
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().conditional.rules).toEqual([rule]);
  });

  it('clears data validation from Home > Clear All with undo support', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'A');
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { bold: true, validation: { kind: 'list', source: ['A', 'B'] } },
    );
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mounted.instance.history.clear();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-ribbon-command="clearFormat"] button')?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="all"]')?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(mounted.instance.store.getState().format.formats.has('0:0:0')).toBe(false);
    expect(mounted.instance.history.undo()).toBe(true);
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'A',
    });
    expect(mounted.instance.store.getState().format.formats.get('0:0:0')?.validation).toEqual({
      kind: 'list',
      source: ['A', 'B'],
    });
  });

  it('fills month series from the Fill menu', async () => {
    mounted = await mountReactSpreadsheet();
    const jan31 = dateSerial(2026, 1, 31);
    mounted.instance.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, jan31);
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'number', value: jan31 },
    );
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const fillButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="fillHome"] button',
    );
    await act(async () => {
      fillButton?.click();
      await flush();
    });
    const months = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="months"]');
    await act(async () => {
      months?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: dateSerial(2026, 2, 28),
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: dateSerial(2026, 3, 31),
    });
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'blank',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'blank',
    });

    mounted.instance.history.redo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: dateSerial(2026, 2, 28),
    });
  });

  it('opens AutoSum menu and inserts the selected aggregate formula', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 20);
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 2, col: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const autoSumButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="autosum"] button',
    );
    await act(async () => {
      autoSumButton?.click();
      await flush();
    });
    const average = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="AVERAGE"]');
    await act(async () => {
      average?.click();
      await flush();
    });

    expect(mounted.instance.workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe(
      '=AVERAGE(A1:A2)',
    );
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'blank',
    });

    mounted.instance.history.redo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe(
      '=AVERAGE(A1:A2)',
    );
  });

  it('opens More Functions from the AutoSum menu', async () => {
    mounted = await mountReactSpreadsheet();
    const fxSpy = vi.spyOn(mounted.instance, 'openFunctionArguments');
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const autoSumButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="autosum"] button',
    );
    await act(async () => {
      autoSumButton?.click();
      await flush();
    });
    const more = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="MORE"]');
    await act(async () => {
      more?.click();
      await flush();
    });

    expect(fxSpy).toHaveBeenCalledWith();
  });

  it('routes Formulas function-library buttons to the matching function argument helper', async () => {
    mounted = await mountReactSpreadsheet();
    const fxSpy = vi.spyOn(mounted.instance, 'openFunctionArguments');
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    const expected = [
      ['fx', undefined, false],
      ['sum', 'SUM', false],
      ['avg', 'AVERAGE', false],
      ['ifFormula', 'IF', true],
      ['xlookupFormula', 'XLOOKUP', true],
      ['concatFormula', 'CONCAT', true],
      ['todayFormula', 'TODAY', true],
      ['pmtFormula', 'PMT', true],
      ['roundFormula', 'ROUND', true],
    ] as const;

    for (const [command, fnName, menu] of expected) {
      const button = toolbar.host.querySelector<HTMLButtonElement>(
        menu ? `[data-ribbon-command="${command}"] button` : `[data-ribbon-command="${command}"]`,
      );
      await act(async () => {
        button?.click();
        await flush();
      });
      if (menu && fnName !== undefined) {
        await act(async () => {
          toolbar.host
            .querySelector<HTMLButtonElement>(
              `[data-ribbon-command="${command}"] [data-cell-action="${fnName}"]`,
            )
            ?.click();
          await flush();
        });
      }
      if (fnName === undefined) expect(fxSpy).toHaveBeenLastCalledWith();
      else expect(fxSpy).toHaveBeenLastCalledWith(fnName);
    }
  });

  it('opens seeded function arguments with the selected range as the first argument', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 });
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    await act(async () => {
      toolbar?.host.querySelector<HTMLButtonElement>('[data-ribbon-command="sum"]')?.click();
      await flush();
    });

    expect(
      document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input')?.value,
    ).toBe('B2:C4');
    expect(document.querySelector<HTMLElement>('.fc-fxdialog__preview')?.textContent).toBe(
      '=SUM(B2:C4)',
    );
  });

  it('routes Find & Select menu items to the matching dialogs', async () => {
    mounted = await mountReactSpreadsheet();
    const findSpy = vi.spyOn(mounted.instance, 'openFindReplace');
    const goToSpy = vi.spyOn(mounted.instance, 'openGoTo');
    const goToSpecialSpy = vi.spyOn(mounted.instance, 'openGoToSpecial');
    const cfRulesSpy = vi.spyOn(mounted.instance, 'openCfRulesDialog');
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const findButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="findHome"] button',
    );
    await act(async () => {
      findButton?.click();
      await flush();
    });
    const replace = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="replace"]');
    await act(async () => {
      replace?.click();
      await flush();
    });
    expect(findSpy).toHaveBeenCalledWith('replace');

    await act(async () => {
      findButton?.click();
      await flush();
    });
    const goTo = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="go-to"]');
    await act(async () => {
      goTo?.click();
      await flush();
    });
    expect(goToSpy).toHaveBeenCalledTimes(1);

    await act(async () => {
      findButton?.click();
      await flush();
    });
    const goToSpecial = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="go-to-special"]',
    );
    await act(async () => {
      goToSpecial?.click();
      await flush();
    });
    expect(goToSpecialSpy).toHaveBeenCalledTimes(1);

    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 1 }, 'cf');
    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 3 }, 'cf');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.addConditionalRule(mounted.instance.store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 },
      color: '#70ad47',
      showValue: true,
    });
    await act(async () => {
      findButton?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-cell-action="conditional-format"]')
        ?.click();
      await flush();
    });
    expect(cfRulesSpy).not.toHaveBeenCalled();
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: { sheet: 0, row: 1, col: 1 },
      range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      extraRanges: [{ sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 }],
    });

    setComment(mounted.instance.store, { sheet: 0, row: 0, col: 0 }, 'first');
    setComment(mounted.instance.store, { sheet: 0, row: 2, col: 2 }, 'second');
    await act(async () => {
      findButton?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="comments"]')?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: { sheet: 0, row: 0, col: 0 },
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      extraRanges: [{ sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }],
    });
  });

  it('selects formulas, constants, and data validation cells from Find & Select', async () => {
    mounted = await mountReactSpreadsheet();
    const formula = { sheet: 0, row: 1, col: 1 };
    const constant = { sheet: 0, row: 3, col: 3 };
    const validated = { sheet: 0, row: 4, col: 4 };
    mounted.instance.workbook.setFormula(formula, '=1+1');
    mounted.instance.workbook.setNumber(constant, 42);
    mounted.instance.workbook.setText(validated, 'Open');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setCellFormat(mounted.instance.store, validated, {
      validation: { kind: 'list', source: ['Open', 'Closed'] },
    });
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    const pickFindAction = async (
      action:
        | 'formulas'
        | 'constants'
        | 'numbers'
        | 'text'
        | 'errors'
        | 'data-validation',
    ): Promise<void> => {
      await act(async () => {
        toolbar.host
          .querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"] button')
          ?.click();
        await flush();
      });
      await act(async () => {
        toolbar.host.querySelector<HTMLButtonElement>(`[data-cell-action="${action}"]`)?.click();
        await flush();
      });
    };

    await pickFindAction('formulas');
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: formula,
      range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
    });

    await pickFindAction('constants');
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: constant,
      range: { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 },
      extraRanges: [{ sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 }],
    });

    await pickFindAction('numbers');
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: formula,
      range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      extraRanges: [{ sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 }],
    });

    const text = { sheet: 0, row: 5, col: 1 };
    const errorText = { sheet: 0, row: 6, col: 1 };
    mounted.instance.workbook.setText(text, 'plain');
    mounted.instance.workbook.setText(errorText, '#DIV/0!');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));

    await pickFindAction('text');
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: validated,
      range: { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
      extraRanges: [{ sheet: 0, r0: 5, c0: 1, r1: 5, c1: 1 }],
    });

    await pickFindAction('errors');
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: errorText,
      range: { sheet: 0, r0: 6, c0: 1, r1: 6, c1: 1 },
    });

    await pickFindAction('data-validation');
    expect(mounted.instance.store.getState().selection).toMatchObject({
      active: validated,
      range: { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
    });
  });

  it('reports when Find & Select cannot find matching cells', async () => {
    mounted = await mountReactSpreadsheet();
    toolbar = await renderToolbar(mounted, { activeTab: 'home', onTabChange: vi.fn() });

    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"] button')
        ?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="formulas"]')?.click();
      await flush();
    });

    expect(toolbar.host.querySelector('.demo__modal')?.textContent).toContain(
      'No matching cells were found.',
    );
  });

  it('opens Defined Names menu actions from the Formulas tab', async () => {
    mounted = await mountReactSpreadsheet();
    const namesSpy = vi.spyOn(mounted.instance, 'openNamedRangeDialog');
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    const namesButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="namedRanges"] button',
    );
    await act(async () => {
      namesButton?.click();
      await flush();
    });
    const defineName = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="define"]');
    await act(async () => {
      defineName?.click();
      await flush();
    });

    expect(namesSpy).toHaveBeenCalledTimes(1);
  });

  it('creates defined names from the selected top row with undo support', async () => {
    mounted = await mountReactSpreadsheet();
    const registry = new Map<string, string>();
    Object.defineProperty(mounted.instance.workbook, 'capabilities', {
      configurable: true,
      value: { ...mounted.instance.workbook.capabilities, definedNameMutate: true },
    });
    vi.spyOn(mounted.instance.workbook, 'definedNames').mockImplementation(function* () {
      for (const [name, formula] of registry) yield { name, formula };
    });
    vi.spyOn(mounted.instance.workbook, 'setDefinedNameEntry').mockImplementation(
      (name, formula) => {
        if (formula) registry.set(name, formula);
        else registry.delete(name);
        return true;
      },
    );
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Sales Total');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, '2026 Rate');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 10);
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 0.08);
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'text', value: 'Sales Total' },
    );
    mutators.setCell(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 1 },
      { kind: 'text', value: '2026 Rate' },
    );
    mounted.instance.history.clear();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    const namesButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="namedRanges"] button',
    );
    await act(async () => {
      namesButton?.click();
      await flush();
    });
    const createTopRow = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="createTopRow"]',
    );
    await act(async () => {
      createTopRow?.click();
      await flush();
    });

    expect([...mounted.instance.workbook.definedNames()]).toEqual([
      { name: 'Sales_Total', formula: '=$A$2:$A$2' },
      { name: '_2026_Rate', formula: '=$B$2:$B$2' },
    ]);
    expect(mounted.instance.history.undo()).toBe(true);
    expect([...mounted.instance.workbook.definedNames()]).toEqual([]);

    mounted.instance.history.clear();
    await act(async () => {
      namesButton?.click();
      await flush();
    });
    const createBottomRow = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="createBottomRow"]',
    );
    await act(async () => {
      createBottomRow?.click();
      await flush();
    });
    expect([...mounted.instance.workbook.definedNames()]).toEqual([
      { name: '_10', formula: '=$A$1:$A$1' },
      { name: '_0.08', formula: '=$B$1:$B$1' },
    ]);
    expect(mounted.instance.history.undo()).toBe(true);
    expect([...mounted.instance.workbook.definedNames()]).toEqual([]);

    mounted.instance.history.clear();
    await act(async () => {
      namesButton?.click();
      await flush();
    });
    const createRightColumn = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="createRightColumn"]',
    );
    await act(async () => {
      createRightColumn?.click();
      await flush();
    });
    expect([...mounted.instance.workbook.definedNames()]).toEqual([
      { name: '_2026_Rate', formula: '=$A$1:$A$1' },
      { name: '_0.08', formula: '=$A$2:$A$2' },
    ]);
    expect(mounted.instance.history.undo()).toBe(true);
    expect([...mounted.instance.workbook.definedNames()]).toEqual([]);
  });

  it('opens Calculation Options menu and writes calc mode metadata', async () => {
    mounted = await mountReactSpreadsheet();
    let calcMode: 0 | 1 | 2 | null = 0;
    vi.spyOn(mounted.instance.workbook, 'calcMode').mockImplementation(() => calcMode);
    const setCalcMode = vi
      .spyOn(mounted.instance.workbook, 'setCalcMode')
      .mockImplementation((mode) => {
        calcMode = mode;
        return true;
      });
    const iterativeSpy = vi.spyOn(mounted.instance, 'openIterativeDialog');
    toolbar = await renderToolbar(mounted, { activeTab: 'formulas', onTabChange: vi.fn() });

    const calcButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="calcOptions"] button',
    );
    await act(async () => {
      calcButton?.click();
      await flush();
    });
    const manual = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="manual"]');
    await act(async () => {
      manual?.click();
      await flush();
    });
    expect(setCalcMode).toHaveBeenCalledWith(1);

    await act(async () => {
      calcButton?.click();
      await flush();
    });
    expect(
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-cell-action="manual"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
    expect(
      toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="manual"]')?.className,
    ).toContain('demo__rb--active');
    expect(
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-cell-action="auto"]')
        ?.getAttribute('aria-checked'),
    ).toBe('false');

    const autoNoTable = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="autoNoTable"]',
    );
    await act(async () => {
      autoNoTable?.click();
      await flush();
    });
    expect(setCalcMode).toHaveBeenLastCalledWith(2);

    await act(async () => {
      calcButton?.click();
      await flush();
    });
    expect(
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-cell-action="autoNoTable"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
    const iterative = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-cell-action="iterative"]',
    );
    await act(async () => {
      iterative?.click();
      await flush();
    });
    expect(iterativeSpy).toHaveBeenCalledTimes(1);
  });

  it('opens Data > Filter menu and clears active filters', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const filterButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="filter"] button',
    );
    await act(async () => {
      filterButton?.click();
      await flush();
    });
    const toggle = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="toggle"]');
    await act(async () => {
      toggle?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().ui.filterRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 3,
      c1: 1,
    });
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().ui.filterRange).toBeNull();
    mounted.instance.history.redo();
    expect(mounted.instance.store.getState().ui.filterRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 3,
      c1: 1,
    });

    await act(async () => {
      filterButton?.click();
      await flush();
    });
    const clear = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="clear"]');
    await act(async () => {
      clear?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().ui.filterRange).toBeNull();

    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().ui.filterRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 3,
      c1: 1,
    });

    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    mounted.instance.history.undo();
    expect(mounted.instance.store.getState().ui.filterRange).toBeNull();
    await act(async () => {
      filterButton?.click();
      await flush();
    });
    const advanced = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="advanced"]');
    await act(async () => {
      advanced?.click();
      await flush();
    });
    expect(toolbar.host.textContent).toContain('Advanced Filter');
    expect(toolbar.host.querySelector<HTMLInputElement>('.demo__modal input')?.value).toBe('A1:B4');
  });

  it('applies Data > Filter > Advanced criteria in place', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Item');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'Qty');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'paper');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 24);
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'ink');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 1 }, 6);
    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 0 }, 'paper');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 3, col: 1 }, 2);
    mounted.instance.workbook.setText({ sheet: 0, row: 5, col: 0 }, 'Item');
    mounted.instance.workbook.setText({ sheet: 0, row: 5, col: 1 }, 'Qty');
    mounted.instance.workbook.setText({ sheet: 0, row: 6, col: 0 }, 'paper');
    mounted.instance.workbook.setText({ sheet: 0, row: 6, col: 1 }, '>10');
    mounted.instance.workbook.setText({ sheet: 0, row: 7, col: 0 }, 'ink');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const filterButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="filter"] button',
    );
    await act(async () => {
      filterButton?.click();
      await flush();
    });
    const advanced = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="advanced"]');
    await act(async () => {
      advanced?.click();
      await flush();
    });
    const inputs = toolbar.host.querySelectorAll<HTMLInputElement>('.demo__modal input');
    await act(async () => {
      if (inputs[1]) {
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
          inputs[1],
          'A6:B8',
        );
        inputs[1].dispatchEvent(new Event('input', { bubbles: true }));
      }
      await flush();
    });
    const ok = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__modal-footer .demo__btn--primary',
    );
    await act(async () => {
      ok?.click();
      await flush();
    });

    const hiddenRows = mounted.instance.store.getState().layout.hiddenRows;
    expect(hiddenRows.has(1)).toBe(false);
    expect(hiddenRows.has(2)).toBe(false);
    expect(hiddenRows.has(3)).toBe(true);
    expect(mounted.instance.store.getState().ui.filterRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 3,
      c1: 1,
    });
    expect(mounted.instance.history.canUndo()).toBe(true);
  });

  it('copies Data > Filter > Advanced results with unique records', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Item');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'Qty');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'paper');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 24);
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'paper');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 1 }, 24);
    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 0 }, 'ink');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 3, col: 1 }, 6);
    mounted.instance.workbook.setText({ sheet: 0, row: 5, col: 0 }, 'Item');
    mounted.instance.workbook.setText({ sheet: 0, row: 6, col: 0 }, 'p*');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const filterButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="filter"] button',
    );
    await act(async () => {
      filterButton?.click();
      await flush();
    });
    const advanced = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="advanced"]');
    await act(async () => {
      advanced?.click();
      await flush();
    });
    const inputs = toolbar.host.querySelectorAll<HTMLInputElement>('.demo__modal input');
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
        inputs[1],
        'A6:A7',
      );
      inputs[1]?.dispatchEvent(new Event('input', { bubbles: true }));
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
        inputs[2],
        'A10',
      );
      inputs[2]?.dispatchEvent(new Event('input', { bubbles: true }));
      inputs[3]?.click();
      await flush();
    });
    const ok = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__modal-footer .demo__btn--primary',
    );
    await act(async () => {
      ok?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 9, col: 0 })).toEqual({
      kind: 'text',
      value: 'Item',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 10, col: 0 })).toEqual({
      kind: 'text',
      value: 'paper',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 11, col: 0 })).toEqual({
      kind: 'blank',
    });
    expect(toolbar.host.textContent).toContain('Copied 2 row(s)');
    expect(mounted.instance.history.canUndo()).toBe(true);
  });

  it('reapplies stored Data > Filter criteria', async () => {
    mounted = await mountReactSpreadsheet();
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Header');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'A');
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'B');
    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 0 }, 'A');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    applyValueFilter(mounted.instance.store.getState(), mounted.instance.store, range, 0, ['B']);
    expect(mounted.instance.store.getState().layout.hiddenRows.has(2)).toBe(true);

    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 0 }, 'B');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.store.getState().layout.hiddenRows.has(3)).toBe(false);
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const filterButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="filter"] button',
    );
    await act(async () => {
      filterButton?.click();
      await flush();
    });
    const reapply = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="reapply"]');
    await act(async () => {
      reapply?.click();
      await flush();
    });

    expect(mounted.instance.store.getState().layout.hiddenRows.has(2)).toBe(true);
    expect(mounted.instance.store.getState().layout.hiddenRows.has(3)).toBe(true);
    expect(mounted.instance.history.canUndo()).toBe(true);
  });

  it('opens Data > Text to Columns menu and splits selected text', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,1');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setCellFormat(
      mounted.instance.store,
      { sheet: 0, row: 0, col: 0 },
      {
        fill: '#c6efce',
      },
    );
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const textToColumnsButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textToColumns"] button',
    );
    await act(async () => {
      textToColumnsButton?.click();
      await flush();
    });
    const comma = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="comma"]');
    await act(async () => {
      comma?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(mounted.instance.store.getState().format.formats.get('0:0:1')).toEqual({
      fill: '#c6efce',
    });
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha,1',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'blank',
    });
    expect(mounted.instance.store.getState().format.formats.has('0:0:1')).toBe(false);

    mounted.instance.history.redo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(mounted.instance.store.getState().format.formats.get('0:0:1')).toEqual({
      fill: '#c6efce',
    });
  });

  it('splits Data > Text to Columns with selected delimiters', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,1;beta gamma');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const textToColumnsButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textToColumns"] button',
    );
    await act(async () => {
      textToColumnsButton?.click();
      await flush();
    });
    const custom = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="custom"]');
    await act(async () => {
      custom?.click();
      await flush();
    });

    const checks = toolbar.host.querySelectorAll<HTMLInputElement>(
      '.demo__modal input[type="checkbox"]',
    );
    await act(async () => {
      checks[2]?.click();
      checks[3]?.click();
      await flush();
    });
    const ok = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__modal-footer .demo__btn--primary',
    );
    await act(async () => {
      ok?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'text',
      value: 'beta',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({
      kind: 'text',
      value: 'gamma',
    });
    expect(mounted.instance.history.canUndo()).toBe(true);
  });

  it('splits Data > Text to Columns while treating consecutive delimiters as one', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,,1,,,beta');
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const textToColumnsButton = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textToColumns"] button',
    );
    await act(async () => {
      textToColumnsButton?.click();
      await flush();
    });
    const custom = toolbar.host.querySelector<HTMLButtonElement>('[data-cell-action="custom"]');
    await act(async () => {
      custom?.click();
      await flush();
    });
    const checks = toolbar.host.querySelectorAll<HTMLInputElement>(
      '.demo__modal input[type="checkbox"]',
    );
    await act(async () => {
      checks[4]?.click();
      await flush();
    });
    const ok = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__modal-footer .demo__btn--primary',
    );
    await act(async () => {
      ok?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'text',
      value: 'beta',
    });
  });

  it('opens Data > Remove Duplicates dialog and honors selected columns', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Name');
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'Value');
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'alpha');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'alpha');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
    mounted.instance.workbook.setText({ sheet: 0, row: 3, col: 0 }, 'beta');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 3, col: 1 }, 3);
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const removeDupes = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="removeDupes"]',
    );
    await act(async () => {
      removeDupes?.click();
      await flush();
    });

    expect(toolbar.host.querySelector('.demo__modal h2')?.textContent).toBe('Remove Duplicates');
    const checks = toolbar.host.querySelectorAll<HTMLInputElement>(
      '.demo__modal input[type="checkbox"]',
    );
    await act(async () => {
      checks[2]?.click();
      await flush();
    });
    const ok = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__modal-footer .demo__btn--primary',
    );
    await act(async () => {
      ok?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'text',
      value: 'beta',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'number',
      value: 3,
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({
      kind: 'blank',
    });
    expect(mounted.instance.history.canUndo()).toBe(true);
  });

  it('runs Data > Outline show detail as show-only, not a hide toggle', async () => {
    mounted = await mountReactSpreadsheet();
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 0 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const group = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="outlineGroup"] button',
    );
    const ungroup = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="outlineUngroup"] button',
    );
    const show = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="outlineShowDetail"]',
    );
    const hide = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="outlineHideDetail"]',
    );

    await act(async () => {
      group?.click();
      await flush();
    });
    expect(toolbar.host.querySelector('[data-ribbon-command="outlineGroup"] [data-cell-action="rows"]')).toBeTruthy();
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="outlineGroup"] [data-cell-action="rows"]')
        ?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.outlineRows.get(1)).toBe(1);

    await act(async () => {
      ungroup?.click();
      await flush();
    });
    await act(async () => {
      toolbar.host
        .querySelector<HTMLButtonElement>('[data-ribbon-command="outlineUngroup"] [data-cell-action="rows"]')
        ?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.outlineRows.size).toBe(0);

    await act(async () => {
      show?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.hiddenRows.size).toBe(0);

    await act(async () => {
      hide?.click();
      await flush();
    });
    expect(Array.from(mounted.instance.store.getState().layout.hiddenRows).sort()).toEqual([
      1, 2, 3,
    ]);

    await act(async () => {
      show?.click();
      await flush();
    });
    expect(mounted.instance.store.getState().layout.hiddenRows.size).toBe(0);
  });

  it('runs Data > Remove Duplicates as one undoable ribbon command', async () => {
    mounted = await mountReactSpreadsheet();
    mounted.instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 1);
    mounted.instance.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'beta');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
    mounted.instance.workbook.setText({ sheet: 0, row: 2, col: 0 }, 'alpha');
    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 1 }, 1);
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    mutators.setRange(mounted.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    toolbar = await renderToolbar(mounted, { activeTab: 'data', onTabChange: vi.fn() });

    const button = toolbar.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="removeDupes"]',
    );
    await act(async () => {
      button?.click();
      await flush();
    });
    const ok = toolbar.host.querySelector<HTMLButtonElement>(
      '.demo__modal-footer .demo__btn--primary',
    );
    await act(async () => {
      ok?.click();
      await flush();
    });

    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'beta',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'blank',
    });
    expect(mounted.instance.history.canUndo()).toBe(true);

    mounted.instance.history.undo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'number',
      value: 1,
    });

    mounted.instance.history.redo();
    mutators.replaceCells(mounted.instance.store, mounted.instance.workbook.cells(0));
    expect(mounted.instance.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'blank',
    });
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
