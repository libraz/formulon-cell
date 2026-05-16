// Title-bar, backstage navigation, autosave, and "title more" menu wiring
// extracted from main.ts. The factory does NOT own the mutable ribbon/backstage
// state — main.ts retains the module-scope `let`s and exposes them via
// getter/setters on the ctx, because there are many external read/write sites
// in main.ts that would otherwise need rewriting.

import {
  findNext,
  mutators,
  type RibbonReportItem,
  type RibbonTab,
  type SpreadsheetInstance,
} from '@libraz/formulon-cell';

import { focusMenuItem, handleMenuKeydown } from '../menu-a11y.js';
import { createMenu, menuButton, menuSeparator } from './menus/general.js';

export interface BackstageTitleShellText {
  save: string;
  saveAs: string;
  autosave: string;
  autosaveOn: string;
  autosaveOff: string;
  comments: string;
  share: string;
  shareReady: string;
}

export interface BackstageTitleCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  shellText: BackstageTitleShellText;
  ribbonRoot: HTMLElement | null;
  titleSearchInput: HTMLInputElement | null;
  autosaveSwitch: HTMLButtonElement | null;
  statusMetric: HTMLElement | null;
  // Host hooks (late-bound where necessary).
  focusSheet: () => void;
  triggerSave: () => void;
  triggerSaveAs: () => Promise<void> | void;
  renderRibbon: () => void;
  refreshAutosave: () => void;
  projectFormatToolbar: () => void;
  showRibbonReport: (title: string, items: readonly RibbonReportItem[]) => void;
  setRibbonDisplayMenuOpen: (open: boolean) => void;
  // State accessors — backed by module-scope `let`s in main.ts so existing
  // external read/write sites continue to work without modification.
  getActiveRibbonTab: () => RibbonTab;
  setActiveRibbonTab: (tab: RibbonTab) => void;
  getRibbonCollapsed: () => boolean;
  setRibbonCollapsed: (collapsed: boolean) => void;
  getBackstageOpen: () => boolean;
  setBackstageOpen: (open: boolean) => void;
  getBackstageReturnTab: () => RibbonTab;
  setBackstageReturnTab: (tab: RibbonTab) => void;
  getAutosaveEnabled: () => boolean;
  setAutosaveEnabled: (enabled: boolean) => void;
}

export interface BackstageTitleApi {
  selectRibbonTab: (tabId: RibbonTab, focusTab?: boolean) => void;
  setRibbonCollapsedExternal: (next: boolean) => void;
  openBackstage: (focus?: 'back' | 'tab') => void;
  closeBackstage: (focusTab?: boolean) => void;
  seedFindDialogQuery: (query: string) => void;
  runTitleSearch: () => void;
  toggleAutosave: () => void;
  closeTitleMoreMenu: (restoreFocus?: boolean) => void;
  openTitleMoreMenu: () => void;
  titleActionButton: (label: string) => HTMLButtonElement | null;
  titleMoreButton: HTMLButtonElement | null;
  titleMoreMenu: HTMLDivElement;
}

export const createBackstageTitle = (ctx: BackstageTitleCtx): BackstageTitleApi => {
  const {
    getInst,
    ribbonLang,
    shellText,
    ribbonRoot,
    titleSearchInput,
    autosaveSwitch,
    statusMetric,
    focusSheet,
    triggerSave,
    triggerSaveAs,
    renderRibbon,
    refreshAutosave,
    projectFormatToolbar,
    showRibbonReport,
    setRibbonDisplayMenuOpen,
    getActiveRibbonTab,
    setActiveRibbonTab,
    getRibbonCollapsed,
    setRibbonCollapsed: setRibbonCollapsedState,
    getBackstageOpen,
    setBackstageOpen,
    getBackstageReturnTab,
    setBackstageReturnTab,
    getAutosaveEnabled,
    setAutosaveEnabled,
  } = ctx;

  const selectRibbonTab = (tabId: RibbonTab, focusTab = false): void => {
    if (!ribbonRoot) return;
    if (tabId === 'file') {
      openBackstage(focusTab ? 'tab' : 'back');
      return;
    }
    setBackstageOpen(false);
    setRibbonDisplayMenuOpen(false);
    setActiveRibbonTab(tabId);
    const activeRibbonTab = getActiveRibbonTab();
    for (const item of ribbonRoot.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')) {
      const isActive = item.dataset.ribbonTab === activeRibbonTab;
      item.classList.toggle('demo__ribbon-tab--active', isActive);
      item.setAttribute('aria-selected', isActive ? 'true' : 'false');
      item.tabIndex = isActive ? 0 : -1;
      if (focusTab && isActive) item.focus({ preventScroll: true });
    }
    for (const panel of ribbonRoot.querySelectorAll<HTMLElement>('[data-ribbon-panel]')) {
      panel.hidden = panel.dataset.ribbonPanel !== activeRibbonTab;
    }
    ribbonRoot.querySelector('.demo__ribbon-display-menu')?.remove();
    ribbonRoot
      .querySelector<HTMLButtonElement>('[data-ribbon-toggle]')
      ?.setAttribute('aria-expanded', 'false');
  };

  const setRibbonCollapsedExternal = (next: boolean): void => {
    setRibbonCollapsedState(next);
    const ribbonCollapsed = getRibbonCollapsed();
    for (const shell of ribbonRoot?.querySelectorAll<HTMLElement>('.demo__ribbon-shell') ?? []) {
      shell.classList.toggle('demo__ribbon-shell--collapsed', ribbonCollapsed);
    }
    for (const tabs of ribbonRoot?.querySelectorAll<HTMLElement>('.demo__ribbon-tabs') ?? []) {
      tabs.dataset.ribbonCollapsed = ribbonCollapsed ? 'true' : 'false';
    }
    for (const item of ribbonRoot?.querySelectorAll<HTMLButtonElement>(
      '[data-ribbon-display-option]',
    ) ?? []) {
      item.setAttribute(
        'aria-checked',
        item.dataset.ribbonDisplayOption === (ribbonCollapsed ? 'collapsed' : 'expanded')
          ? 'true'
          : 'false',
      );
    }
  };

  const openBackstage = (focus: 'back' | 'tab' = 'back'): void => {
    if (!getBackstageOpen() && getActiveRibbonTab() !== 'file')
      setBackstageReturnTab(getActiveRibbonTab());
    setActiveRibbonTab('file');
    setBackstageOpen(true);
    setRibbonCollapsedState(false);
    setRibbonDisplayMenuOpen(false);
    renderRibbon();
    const selector =
      focus === 'tab' ? '[data-ribbon-tab="file"]' : '[data-backstage-action="back"]';
    ribbonRoot?.querySelector<HTMLButtonElement>(selector)?.focus({ preventScroll: true });
  };

  const closeBackstage = (focusTab = false): void => {
    if (!getBackstageOpen()) return;
    setBackstageOpen(false);
    setActiveRibbonTab(getBackstageReturnTab());
    renderRibbon();
    if (focusTab) {
      ribbonRoot
        ?.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${getActiveRibbonTab()}"]`)
        ?.focus({ preventScroll: true });
    }
  };

  const titleActionButton = (label: string): HTMLButtonElement | null =>
    document.querySelector<HTMLButtonElement>(`.app__title [data-shell-i18n-label="${label}"]`);

  titleActionButton('home')?.addEventListener('click', () => {
    closeBackstage();
    selectRibbonTab('home', true);
  });

  titleActionButton('save')?.addEventListener('click', () => {
    triggerSave();
  });

  titleActionButton('saveAs')?.addEventListener('click', () => {
    void triggerSaveAs();
  });

  titleActionButton('undo')?.addEventListener('click', () => {
    const inst = getInst();
    if (inst?.undo()) focusSheet();
  });

  titleActionButton('redo')?.addEventListener('click', () => {
    const inst = getInst();
    if (inst?.redo()) focusSheet();
  });

  titleActionButton('comments')?.addEventListener('click', () => {
    getInst()?.openCommentDialog();
  });

  titleActionButton('share')?.addEventListener('click', () => {
    showRibbonReport(shellText.share, [
      { severity: 'info', label: shellText.share, detail: shellText.shareReady },
    ]);
  });

  const seedFindDialogQuery = (query: string): void => {
    requestAnimationFrame(() => {
      const input = document.querySelector<HTMLInputElement>('.fc-find input[type="text"]');
      if (!input) return;
      input.value = query;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.focus();
      input.select();
    });
  };

  const runTitleSearch = (): void => {
    const query = titleSearchInput?.value.trim() ?? '';
    const inst = getInst();
    if (!query || !inst) return;
    const state = inst.store.getState();
    const match = findNext(
      state,
      { query, within: 'sheet', searchBy: 'rows', lookIn: 'values' },
      state.selection.active,
      'next',
    );
    inst.openFindReplace('find');
    seedFindDialogQuery(query);
    if (match) {
      mutators.setActive(inst.store, match.addr);
      projectFormatToolbar();
    } else if (statusMetric) {
      statusMetric.textContent =
        ribbonLang === 'ja' ? `「${query}」は見つかりませんでした` : `No matches for "${query}"`;
    }
  };

  titleSearchInput?.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    runTitleSearch();
  });

  document.addEventListener('keydown', (event) => {
    if (event.key.toLowerCase() !== 'u' || !event.metaKey || !event.ctrlKey) return;
    event.preventDefault();
    titleSearchInput?.focus();
    titleSearchInput?.select();
  });

  const toggleAutosave = (): void => {
    setAutosaveEnabled(!getAutosaveEnabled());
    refreshAutosave();
    if (statusMetric)
      statusMetric.textContent = getAutosaveEnabled()
        ? shellText.autosaveOn
        : shellText.autosaveOff;
  };

  autosaveSwitch?.addEventListener('click', toggleAutosave);

  const titleMoreButton = titleActionButton('more');
  const titleMoreMenu = createMenu('menu-title-more');
  titleMoreMenu.classList.add('app__title-more-menu');
  titleMoreMenu.append(
    menuButton(shellText.save, 'titleMoreAction', 'save'),
    menuButton(shellText.saveAs, 'titleMoreAction', 'save-as'),
    menuButton(shellText.autosave, 'titleMoreAction', 'autosave'),
    menuSeparator(),
    menuButton(shellText.comments, 'titleMoreAction', 'comments'),
    menuButton(shellText.share, 'titleMoreAction', 'share'),
  );
  document.body.appendChild(titleMoreMenu);

  const closeTitleMoreMenu = (restoreFocus = false): void => {
    titleMoreMenu.hidden = true;
    titleMoreButton?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) titleMoreButton?.focus({ preventScroll: true });
  };

  const openTitleMoreMenu = (): void => {
    if (!titleMoreButton) return;
    const rect = titleMoreButton.getBoundingClientRect();
    titleMoreMenu.style.left = `${Math.round(rect.left)}px`;
    titleMoreMenu.style.top = `${Math.round(rect.bottom + 4)}px`;
    titleMoreMenu.hidden = false;
    titleMoreButton.setAttribute('aria-haspopup', 'menu');
    titleMoreButton.setAttribute('aria-expanded', 'true');
    focusMenuItem(titleMoreMenu, 'first');
  };

  titleMoreButton?.addEventListener('click', () => {
    if (titleMoreMenu.hidden) openTitleMoreMenu();
    else closeTitleMoreMenu(true);
  });

  titleMoreMenu.addEventListener('click', (event) => {
    const item = (event.target as Element | null)?.closest<HTMLButtonElement>(
      '[data-title-more-action]',
    );
    const action = item?.dataset.titleMoreAction;
    if (!action) return;
    closeTitleMoreMenu();
    if (action === 'save') triggerSave();
    else if (action === 'save-as') void triggerSaveAs();
    else if (action === 'autosave') toggleAutosave();
    else if (action === 'comments') getInst()?.openCommentDialog();
    else if (action === 'share') {
      showRibbonReport(shellText.share, [
        { severity: 'info', label: shellText.share, detail: shellText.shareReady },
      ]);
    }
  });

  titleMoreMenu.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, titleMoreMenu, {
      close: closeTitleMoreMenu,
      restoreFocusTo: titleMoreButton ?? undefined,
    });
  });

  document.addEventListener('pointerdown', (event) => {
    if (titleMoreMenu.hidden) return;
    const target = event.target as Element | null;
    if (titleMoreMenu.contains(target)) return;
    if (titleMoreButton?.contains(target)) return;
    closeTitleMoreMenu();
  });

  return {
    selectRibbonTab,
    setRibbonCollapsedExternal,
    openBackstage,
    closeBackstage,
    seedFindDialogQuery,
    runTitleSearch,
    toggleAutosave,
    closeTitleMoreMenu,
    openTitleMoreMenu,
    titleActionButton,
    titleMoreButton,
    titleMoreMenu,
  };
};
