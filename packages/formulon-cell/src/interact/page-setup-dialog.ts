// Page Setup dialog. Lets the user edit the active sheet's `PageSetup` —
// orientation, paper size, margins (inches), header / footer slots, print
// titles, scale, gridlines / headings toggles. OK pushes the resulting patch
// through `mutators.setPageSetup` wrapped in a single history entry so undo
// reverts the whole apply atomically.
import type { History } from '../commands/history.js';
import { printableMarginAdjustments } from '../commands/print.js';
import {
  normalizePrintableBounds,
  normalizePrinterProfileId,
  normalizePrinterProfiles,
  type PrinterProfile,
} from '../commands/printer-profile.js';
import { recordPageSetupChange } from '../commands/slice-history.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  defaultPageSetup,
  getPageSetup,
  mutators,
  type PageMargins,
  type PageSetup,
  type SpreadsheetStore,
} from '../store/store.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogIconButton,
  appendDialogTabPair,
  createDialogShell,
} from './dialog-shell.js';
import { createHeaderFooterTab } from './page-setup-dialog-tabs/header-footer.js';
import { createMarginsTab } from './page-setup-dialog-tabs/margins.js';
import { createPageTab } from './page-setup-dialog-tabs/page.js';
import { createSheetTab } from './page-setup-dialog-tabs/sheet.js';

export interface PageSetupDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Shared history. When provided, OK pushes one page-setup snapshot entry. */
  history?: History | null;
  /** Optional host hook for resolving printer non-printable bounds after the
   *  user changes paper size or orientation. Browser APIs do not expose these
   *  values, so Electron/native hosts can refresh their profile here. */
  resolvePrintableBounds?: (
    setup: PageSetup,
    sheet: number,
    previous: PageSetup,
    printerProfileId: string | undefined,
  ) => Partial<PageMargins> | null | undefined;
  getPrinterProfiles?: () => readonly PrinterProfile[] | undefined;
  getPrinterProfileId?: () => string | undefined;
  setPrinterProfileId?: (next: string | undefined) => void;
  refreshPrinterProfiles?: () =>
    | readonly PrinterProfile[]
    | undefined
    | Promise<readonly PrinterProfile[] | undefined>;
}

export interface PageSetupDialogHandle {
  open(tab?: PageSetupDialogTab): void;
  close(): void;
  detach(): void;
}

export type PageSetupDialogTab = 'page' | 'margins' | 'headerFooter' | 'sheet';

export function attachPageSetupDialog(deps: PageSetupDialogDeps): PageSetupDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.pageSetup;
  let openingPrinterProfileId = normalizePrinterProfileId(deps.getPrinterProfileId?.());

  const shell = createDialogShell({
    host,
    className: 'fc-pgsetup',
    ariaLabel: t.title,
    onDismiss: () => onCancel(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-pgsetup__panel');
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  const headerTitle = document.createElement('span');
  headerTitle.textContent = t.title;
  // Distinct from the footer Cancel button — otherwise Playwright's strict
  // `getByRole('button', { name: 'Cancel' })` finds both and throws.
  // Heuristic: if t.cancel is non-ASCII (likely Japanese) use the Japanese
  // "Close" — otherwise English.
  const closeLabel = Array.from(t.cancel).some((ch) => ch.charCodeAt(0) > 0x7f)
    ? '閉じる'
    : 'Close';
  header.appendChild(headerTitle);
  const headerCloseBtn = appendDialogIconButton(header, {
    label: '',
    ariaLabel: closeLabel,
    baseClass: 'fc-fmtdlg__close',
  });
  panel.appendChild(header);

  const tabsStrip = document.createElement('div');
  tabsStrip.className = 'fc-fmtdlg__tabs';
  tabsStrip.setAttribute('role', 'tablist');
  tabsStrip.setAttribute('aria-label', t.title);
  panel.appendChild(tabsStrip);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  const tabDefs: { id: PageSetupDialogTab; label: string }[] = [
    { id: 'page', label: t.tabPage },
    { id: 'margins', label: t.tabMargins },
    { id: 'headerFooter', label: t.tabHeaderFooter },
    { id: 'sheet', label: t.tabSheet },
  ];
  const tabButtons = new Map<PageSetupDialogTab, HTMLButtonElement>();
  const tabPanels = new Map<PageSetupDialogTab, HTMLDivElement>();
  for (const def of tabDefs) {
    const { button, panel: tabPanel } = appendDialogTabPair(tabsStrip, body, {
      id: def.id,
      label: def.label,
      tabId: `fc-pgsetup-tab-${def.id}`,
      panelId: `fc-pgsetup-panel-${def.id}`,
      panelClass: 'fc-fmtdlg__panel-tab fc-pgsetup__tab-panel',
      tabDatasetKey: 'pgsetupTab',
      panelDatasetKey: 'pgsetupTab',
    });
    tabButtons.set(def.id, button);
    tabPanels.set(def.id, tabPanel);
  }

  const pagePanel = tabPanels.get('page') as HTMLDivElement;
  const marginsPanel = tabPanels.get('margins') as HTMLDivElement;
  const headerFooterPanel = tabPanels.get('headerFooter') as HTMLDivElement;
  const sheetPanel = tabPanels.get('sheet') as HTMLDivElement;
  const tabCtx = { t, strings, store, on: shell.on };
  const pageTab = createPageTab(pagePanel, tabCtx);
  const { printerRow, printerSelect, printerRefreshBtn, printerStatus } = pageTab;
  printerRefreshBtn.hidden = !deps.refreshPrinterProfiles;
  const marginsTab = createMarginsTab(marginsPanel, tabCtx);
  const headerFooterTab = createHeaderFooterTab(headerFooterPanel, tabCtx);
  const sheetTab = createSheetTab(sheetPanel, tabCtx, () => setActiveTab('sheet'));

  // ── Footer / buttons ────────────────────────────────────────────────────
  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
  });
  panel.appendChild(footer);

  /** Snapshot of the dialog values when it opened. Used by Cancel to revert
   *  inline edits and (more importantly) by OK to push a single history entry
   *  spanning the whole apply. */
  let opening: PageSetup = defaultPageSetup();
  let activeTab: PageSetupDialogTab = 'page';

  const setActiveTab = (id: PageSetupDialogTab): void => {
    activeTab = id;
    for (const [tabId, btn] of tabButtons) {
      btn.setAttribute('aria-selected', tabId === id ? 'true' : 'false');
      btn.tabIndex = tabId === id ? 0 : -1;
    }
    for (const [tabId, tabPanel] of tabPanels) {
      tabPanel.hidden = tabId !== id;
    }
  };

  const tabOrder = Array.from(tabButtons.keys());
  const focusTabByIndex = (index: number): void => {
    const next = tabOrder[(index + tabOrder.length) % tabOrder.length];
    if (!next) return;
    setActiveTab(next);
    tabButtons.get(next)?.focus();
  };

  const onTabClick = (event: MouseEvent): void => {
    const btn = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-pgsetup-tab]');
    const id = btn?.dataset.pgsetupTab as PageSetupDialogTab | undefined;
    if (!btn || !id) return;
    setActiveTab(id);
    btn.focus();
  };

  const onTabKeyDown = (event: KeyboardEvent): void => {
    const btn = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-pgsetup-tab]');
    const id = btn?.dataset.pgsetupTab as PageSetupDialogTab | undefined;
    const index = id ? tabOrder.indexOf(id) : -1;
    if (index < 0) return;
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      event.preventDefault();
      focusTabByIndex(index + 1);
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      event.preventDefault();
      focusTabByIndex(index - 1);
    } else if (event.key === 'Home') {
      event.preventDefault();
      focusTabByIndex(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      focusTabByIndex(tabOrder.length - 1);
    }
  };

  const renderPrinterProfiles = (nextProfiles?: readonly PrinterProfile[]): void => {
    const profiles = deps.getPrinterProfiles?.() ?? [];
    const effectiveProfiles = normalizePrinterProfiles(nextProfiles ?? profiles) ?? [];
    const selected = normalizePrinterProfileId(deps.getPrinterProfileId?.()) ?? '';
    printerSelect.replaceChildren();
    appendDialogSelectOptions(printerSelect, [
      { value: '', label: t.printerProfileAutomatic },
      ...effectiveProfiles.flatMap((profile) =>
        profile.id ? [{ value: profile.id, label: profile.name || profile.id }] : [],
      ),
    ]);
    printerRow.hidden = effectiveProfiles.length === 0 && !deps.refreshPrinterProfiles;
    printerSelect.value = selected;
    if (printerSelect.value !== selected) printerSelect.value = '';
  };

  const refreshPrinterProfilesFromHost = async (): Promise<void> => {
    if (!deps.refreshPrinterProfiles) return;
    projectDisabledState(printerRefreshBtn, true, t.printerProfileRefreshInProgress, {
      datasetKey: 'disabledReason',
      titlePrefix: t.printerProfileRefresh,
    });
    printerStatus.textContent = '';
    try {
      const profiles = await deps.refreshPrinterProfiles();
      renderPrinterProfiles(profiles);
    } catch {
      printerStatus.textContent = t.printerProfileRefreshFailed;
    } finally {
      projectDisabledState(printerRefreshBtn, false, null, {
        datasetKey: 'disabledReason',
        titlePrefix: t.printerProfileRefresh,
      });
    }
  };

  const hydrateFrom = (setup: PageSetup): void => {
    openingPrinterProfileId = normalizePrinterProfileId(deps.getPrinterProfileId?.());
    renderPrinterProfiles();
    opening = { ...setup, margins: { ...setup.margins } };
    pageTab.hydrate(setup);
    marginsTab.hydrate(setup);
    headerFooterTab.hydrate(setup);
    sheetTab.hydrate(setup);
    setActiveTab('page');
  };

  const collectFromInputs = (): PageSetup => ({
    ...pageTab.collect(),
    ...marginsTab.collect(),
    ...headerFooterTab.collect(),
    ...sheetTab.collect(),
  });

  const updatePrintableWarning = (): void => {
    marginsTab.renderWarning(printableMarginAdjustments(collectFromInputs()));
  };

  const onOk = (): void => {
    if (!sheetTab.validate()) return;
    const sheet = store.getState().data.sheetIndex;
    const next = collectFromInputs();
    const nextPrinterProfileId = printerSelect.value || undefined;
    const paperChanged =
      next.paperSize !== opening.paperSize || next.orientation !== opening.orientation;
    const printerProfileChanged = nextPrinterProfileId !== openingPrinterProfileId;
    if (paperChanged || printerProfileChanged) {
      const resolved = deps.resolvePrintableBounds?.(next, sheet, opening, nextPrinterProfileId);
      if (resolved !== undefined) {
        next.printableBounds = resolved ? normalizePrintableBounds(resolved) : undefined;
      }
    }
    if (printerProfileChanged) deps.setPrinterProfileId?.(nextPrinterProfileId);
    recordPageSetupChange(history, store, () => {
      mutators.setPageSetup(store, sheet, next);
    });
    api.close();
  };

  const onCancel = (): void => {
    // Cancel: rehydrate inputs so a follow-up open() doesn't show stale text,
    // and skip the slice mutation entirely. The opening snapshot is still in
    // the slice unchanged.
    hydrateFrom(opening);
    api.close();
  };

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
      return;
    }
    if (e.key === 'Enter') {
      if ((e.target as HTMLElement).tagName === 'BUTTON') return;
      // Enter inside a textbox triggers OK — spreadsheet parity.
      e.preventDefault();
      onOk();
    }
  };

  shell.on(tabsStrip, 'click', onTabClick as EventListener);
  shell.on(tabsStrip, 'keydown', onTabKeyDown as EventListener);
  shell.on(headerCloseBtn, 'click', onCancel);
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(printerRefreshBtn, 'click', () => {
    void refreshPrinterProfilesFromHost();
  });
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);
  for (const input of marginsTab.marginInputs) {
    shell.on(input, 'input', updatePrintableWarning);
  }

  const api: PageSetupDialogHandle = {
    open(tab: PageSetupDialogTab = 'page'): void {
      const sheet = store.getState().data.sheetIndex;
      hydrateFrom(getPageSetup(store.getState(), sheet));
      setActiveTab(tab);
      updatePrintableWarning();
      shell.open();
      requestAnimationFrame(() => {
        tabButtons.get(activeTab)?.focus();
      });
    },
    close(): void {
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };

  return api;
}
