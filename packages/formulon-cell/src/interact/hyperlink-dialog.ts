import type { History } from '../commands/history.js';
import { setHyperlinkAction } from '../commands/hyperlink-actions.js';
import { hyperlinkAt } from '../commands/hyperlinks.js';
import { runAfterAllowedRangeUnlock } from '../commands/protection.js';
import { quoteSheetName } from '../engine/address.js';
import { addrKey, type WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { showPrompt } from '../toolbar/dialogs/prompt.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { type DialogEntry, INSTANCE_API_HYPERLINK_ENTRY } from './dialog-entry.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  clearDialogError,
  createDialogButton,
  createDialogShell,
  deferInitialFocus,
  focusAndSelectInput,
  showDialogError,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

export interface HyperlinkDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  history?: History | null;
  /** Workbook getter — lazy so the dialog stays in lockstep with `setWorkbook`
   *  swaps. When the engine supports hyperlinks the OK/Remove path also
   *  flushes the change to the engine for xlsx round-trip. */
  getWb?: () => WorkbookHandle | null;
}

export interface HyperlinkDialogHandle {
  /** `entry` names the entry point OK and Remove are authorized under; the instance API when absent. */
  open(entry?: DialogEntry): void;
  close(): void;
  detach(): void;
}

/**
 * Insert Hyperlink dialog. Edits the hyperlink field on the currently active
 * cell only; multi-cell apply goes through the full Format Cells dialog.
 */
export function attachHyperlinkDialog(deps: HyperlinkDialogDeps): HyperlinkDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  let dialogEntry: DialogEntry = INSTANCE_API_HYPERLINK_ENTRY;
  const strings = deps.strings ?? defaultStrings;
  // The Mac dialog adds display text, a ScreenTip and Web / document / e-mail tabs.
  const mac = isMacHost(host);
  const m = mac ? strings.mac.hyperlinkDialog : null;
  const t = m ? { ...strings.hyperlinkDialog, ...m } : strings.hyperlinkDialog;

  const shell = createDialogShell({
    host,
    className: 'fc-hldlg',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { overlay } = shell;
  const { body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-hldlg__panel'],
  });

  const textInput = (): HTMLInputElement => {
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'fc-fmtdlg__input';
    input.autocomplete = 'off';
    input.spellcheck = false;
    return input;
  };
  /** A caption + control row; the Mac panels use it for every field. */
  const field = (caption: string, control: HTMLElement, className = 'fc-hldlg__field') => {
    const label = document.createElement('label');
    label.className = className;
    const span = document.createElement('span');
    span.textContent = caption;
    label.append(span, control);
    return label;
  };
  const listbox = (className: string): HTMLDivElement => {
    const list = document.createElement('div');
    list.className = className;
    list.setAttribute('role', 'listbox');
    return list;
  };

  const urlInput = document.createElement('input');
  urlInput.type = mac ? 'text' : 'url';
  urlInput.className = 'fc-fmtdlg__input';
  urlInput.placeholder = mac ? '' : t.urlPlaceholder;
  urlInput.autocomplete = 'off';
  urlInput.spellcheck = false;

  type LinkTab = 'web' | 'document' | 'email';
  const displayInput = textInput();
  const tipBtn = createDialogButton({ label: '', baseClass: 'fc-fmtdlg__btn fc-hldlg__tip' });
  const refInput = textInput();
  const places = listbox('fc-hldlg__places');
  const emailInput = textInput();
  const subjectInput = textInput();
  const recentList = listbox('fc-hldlg__recent');
  const tabs = new Map<LinkTab, { tab: HTMLButtonElement; panel: HTMLDivElement }>();
  if (m) {
    const top = document.createElement('div');
    top.className = 'fc-hldlg__top';
    tipBtn.textContent = m.screenTip;
    top.append(field(m.displayText, displayInput, 'fc-hldlg__display'), tipBtn);
    const tablist = document.createElement('div');
    tablist.className = 'fc-hldlg__tabs';
    tablist.setAttribute('role', 'tablist');
    const panelFor = (name: LinkTab, caption: string): HTMLDivElement => {
      const tab = createDialogButton({ label: caption, baseClass: 'fc-hldlg__tab' });
      tab.setAttribute('role', 'tab');
      tab.dataset.tab = name;
      tablist.appendChild(tab);
      const panel = document.createElement('div');
      panel.className = `fc-hldlg__panel-${name} fc-hldlg__${name}`;
      panel.setAttribute('role', 'tabpanel');
      tabs.set(name, { tab, panel });
      return panel;
    };
    const web = panelFor('web', m.tabWeb);
    const intro = document.createElement('p');
    intro.className = 'fc-hldlg__intro';
    intro.textContent = m.webIntro;
    const browse = createDialogButton({
      label: m.browse,
      baseClass: 'fc-fmtdlg__btn fc-hldlg__browse',
    });
    projectDisabledState(browse, true, m.browseUnavailable, { datasetKey: 'disabledReason' });
    const addressRow = document.createElement('div');
    addressRow.className = 'fc-hldlg__address';
    addressRow.append(field(t.url, urlInput), browse);
    web.append(intro, addressRow);
    const doc = panelFor('document', m.tabDocument);
    const placeCaption = document.createElement('div');
    placeCaption.textContent = m.documentPlace;
    doc.append(field(m.cellReference, refInput), placeCaption, places);
    const email = panelFor('email', m.tabEmail);
    const recentCaption = document.createElement('div');
    recentCaption.textContent = m.recentEmails;
    email.append(
      field(m.emailAddress, emailInput),
      field(m.subject, subjectInput),
      recentCaption,
      recentList,
    );
    body.append(top, tablist, ...[...tabs.values()].map(({ panel }) => panel));
  } else {
    const urlRow = document.createElement('div');
    urlRow.className = 'fc-fmtdlg__row';
    const urlLabel = document.createElement('label');
    urlLabel.textContent = t.url;
    urlLabel.appendChild(urlInput);
    urlRow.appendChild(urlLabel);
    body.appendChild(urlRow);
  }

  const errorRow = document.createElement('div');
  errorRow.className = 'fc-hldlg__error';
  errorRow.setAttribute('role', 'alert');
  errorRow.hidden = true;
  body.appendChild(errorRow);

  const removeBtn = appendDialogButton(footer, { label: t.remove });
  // Anchored on the left side so it doesn't get confused with OK/Cancel.
  removeBtn.style.marginRight = 'auto';
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
  });

  const showError = (msg: string): void => {
    showDialogError(errorRow, msg);
  };
  const clearError = (): void => {
    clearDialogError(errorRow);
  };

  const writeHyperlink = (
    next: string | undefined,
    options: { display?: string; tooltip?: string } = {},
  ): void => {
    const state = store.getState();
    const addr = state.selection.active;
    const wb = getWb() ?? undefined;
    const cell = { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col };
    runAfterAllowedRangeUnlock(store, [cell], () => {
      setHyperlinkAction(
        {
          store,
          history,
          ...(wb ? { workbook: wb } : {}),
          origin: dialogEntry.origin,
          commandId: dialogEntry.commandId,
        },
        addr,
        next,
        options,
      );
    });
  };

  let activeTab: LinkTab = 'web';
  let placeSheet: string | null = null;
  let tooltip: string | undefined;
  const recentEmails: string[] = [];
  const unquoteSheet = (name: string): string =>
    name.startsWith("'") && name.endsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
  /** The link the active Mac tab describes; the plain dialog uses the address field. */
  const target = (): string => {
    if (!mac || activeTab === 'web') return urlInput.value.trim();
    if (activeTab === 'document') {
      const ref = refInput.value.trim();
      if (!ref) return '';
      const active = store.getState().selection.active;
      const sheet = placeSheet ?? getWb()?.sheetName(active.sheet) ?? '';
      return sheet ? `#${quoteSheetName(sheet)}!${ref}` : `#${ref}`;
    }
    const address = emailInput.value.trim();
    if (!address) return '';
    const subject = subjectInput.value.trim();
    return `mailto:${address}${subject ? `?subject=${encodeURIComponent(subject)}` : ''}`;
  };

  // The Mac dialog keeps OK disabled until the active tab names a target.
  const syncOk = (): void => {
    if (!mac) return;
    projectDisabledState(okBtn, target() === '', t.errorEmptyUrl);
  };
  for (const input of [urlInput, refInput, emailInput, subjectInput]) {
    input.addEventListener('input', syncOk);
  }

  const selectOption = (list: HTMLElement, option: HTMLElement | null): void => {
    for (const item of list.querySelectorAll<HTMLElement>('[role="option"]')) {
      item.setAttribute('aria-selected', String(item === option));
    }
  };
  const renderOptions = (
    list: HTMLElement,
    values: readonly string[],
    key: string,
    onPick: (value: string, option: HTMLElement) => void,
  ): void => {
    list.replaceChildren(
      ...values.map((value) => {
        const option = document.createElement('div');
        option.setAttribute('role', 'option');
        option.setAttribute('aria-selected', 'false');
        option.dataset[key] = value;
        option.textContent = value;
        option.addEventListener('click', () => onPick(value, option));
        return option;
      }),
    );
  };
  const selectTab = (name: LinkTab): void => {
    activeTab = name;
    for (const [key, { tab, panel }] of tabs) {
      tab.setAttribute('aria-selected', String(key === name));
      tab.tabIndex = key === name ? 0 : -1;
      panel.hidden = key !== name;
    }
    syncOk();
  };
  for (const [name, { tab }] of tabs) tab.addEventListener('click', () => selectTab(name));
  tipBtn.addEventListener('click', () => {
    if (!m) return;
    void showPrompt({
      title: m.screenTipTitle,
      label: m.screenTipLabel,
      initial: tooltip ?? '',
      okLabel: t.ok,
      cancelLabel: t.cancel,
    }).then((value) => {
      if (value !== null) tooltip = value.trim() || undefined;
    });
  });

  const onOk = (): void => {
    if (mac && okBtn.disabled) return;
    const url = target();
    if (!url) {
      showError(t.errorEmptyUrl);
      focusAndSelectInput(urlInput);
      return;
    }
    clearError();
    if (!mac) {
      writeHyperlink(url);
      api.close();
      return;
    }
    const display = displayInput.disabled ? '' : displayInput.value.trim();
    writeHyperlink(url, {
      ...(display ? { display } : {}),
      ...(tooltip ? { tooltip } : {}),
    });
    if (activeTab === 'email') {
      const address = emailInput.value.trim();
      const at = recentEmails.indexOf(address);
      if (at >= 0) recentEmails.splice(at, 1);
      recentEmails.unshift(address);
    }
    api.close();
  };

  const onRemove = (): void => {
    clearError();
    writeHyperlink(undefined);
    api.close();
  };

  const onCancel = (): void => api.close();

  const onOverlayKey = (e: KeyboardEvent): void => {
    if (!isSubmitEnter(e)) return;
    e.preventDefault();
    e.stopPropagation();
    onOk();
  };

  shell.on(okBtn, 'click', onOk);
  shell.on(removeBtn, 'click', onRemove);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  /** Seed the Mac fields from the active cell and its current link. */
  const openMac = (current: string): void => {
    if (!m) return;
    const state = store.getState();
    const { range, active } = state.selection;
    const fmt = state.format.formats.get(addrKey(active));
    tooltip = fmt?.hyperlinkTooltip;
    const single = range.r0 === range.r1 && range.c0 === range.c1;
    const value = state.data.cells.get(addrKey(active))?.value;
    const cellText =
      value?.kind === 'text' ? value.value : value?.kind === 'number' ? String(value.value) : '';
    displayInput.value = single ? (fmt?.hyperlinkDisplay ?? cellText) : m.selectionInDocument;
    projectDisabledState(displayInput, !single, null);
    const wb = getWb();
    const sheets = wb ? Array.from({ length: wb.sheetCount }, (_, i) => wb.sheetName(i)) : [];
    renderOptions(places, sheets, 'sheet', (name, option) => {
      placeSheet = name;
      selectOption(places, option);
      syncOk();
    });
    renderOptions(recentList, recentEmails, 'email', (address, option) => {
      emailInput.value = address;
      selectOption(recentList, option);
      syncOk();
    });
    placeSheet = null;
    refInput.value = '';
    emailInput.value = '';
    subjectInput.value = '';
    if (current.startsWith('#')) {
      const place = current.slice(1);
      const bang = place.lastIndexOf('!');
      placeSheet = bang >= 0 ? unquoteSheet(place.slice(0, bang)) : null;
      refInput.value = bang >= 0 ? place.slice(bang + 1) : place;
      selectOption(places, places.querySelector(`[data-sheet="${CSS.escape(placeSheet ?? '')}"]`));
      urlInput.value = '';
      selectTab('document');
    } else if (current.toLowerCase().startsWith('mailto:')) {
      const [address = '', query = ''] = current.slice(7).split('?');
      emailInput.value = decodeURIComponent(address);
      subjectInput.value = decodeURIComponent(new URLSearchParams(query).get('subject') ?? '');
      urlInput.value = '';
      selectTab('email');
    } else {
      selectTab('web');
    }
  };

  const api: HyperlinkDialogHandle = {
    open(entry = INSTANCE_API_HYPERLINK_ENTRY): void {
      dialogEntry = entry;
      const state = store.getState();
      const addr = state.selection.active;
      const current = hyperlinkAt(state, addr) ?? '';
      urlInput.value = current;
      if (mac) openMac(current);
      syncOk();
      // Hide Remove when there's nothing to remove.
      removeBtn.hidden = !current;
      clearError();
      shell.open();
      deferInitialFocus(overlay, () => {
        focusAndSelectInput(
          activeTab === 'document' ? refInput : activeTab === 'email' ? emailInput : urlInput,
        );
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
