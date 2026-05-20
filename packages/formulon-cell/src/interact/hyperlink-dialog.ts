import { type History, recordFormatChange } from '../commands/history.js';
import { clearHyperlink, hyperlinkAt, setHyperlink } from '../commands/hyperlinks.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  clearDialogError,
  createDialogShell,
  focusAndSelectInput,
  showDialogError,
} from './dialog-shell.js';

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
  open(): void;
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
  const strings = deps.strings ?? defaultStrings;
  const t = strings.hyperlinkDialog;

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

  const urlRow = document.createElement('div');
  urlRow.className = 'fc-fmtdlg__row';
  const urlLabel = document.createElement('label');
  urlLabel.textContent = t.url;
  const urlInput = document.createElement('input');
  urlInput.type = 'url';
  urlInput.className = 'fc-fmtdlg__input';
  urlInput.placeholder = t.urlPlaceholder;
  urlInput.autocomplete = 'off';
  urlInput.spellcheck = false;
  urlLabel.appendChild(urlInput);
  urlRow.appendChild(urlLabel);
  body.appendChild(urlRow);

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

  const writeHyperlink = (next: string | undefined): void => {
    const state = store.getState();
    const addr = state.selection.active;
    const wb = getWb() ?? undefined;
    recordFormatChange(history, store, () => {
      if (next) setHyperlink(store, addr, next, wb);
      else clearHyperlink(store, addr, wb);
    });
  };

  const onOk = (): void => {
    const url = urlInput.value.trim();
    if (!url) {
      showError(t.errorEmptyUrl);
      focusAndSelectInput(urlInput);
      return;
    }
    clearError();
    writeHyperlink(url);
    api.close();
  };

  const onRemove = (): void => {
    clearError();
    writeHyperlink(undefined);
    api.close();
  };

  const onCancel = (): void => api.close();

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
      return;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      onOk();
    }
  };

  shell.on(okBtn, 'click', onOk);
  shell.on(removeBtn, 'click', onRemove);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  const api: HyperlinkDialogHandle = {
    open(): void {
      const state = store.getState();
      const addr = state.selection.active;
      const current = hyperlinkAt(state, addr) ?? '';
      urlInput.value = current;
      // Hide Remove when there's nothing to remove.
      removeBtn.hidden = !current;
      clearError();
      shell.open();
      requestAnimationFrame(() => {
        focusAndSelectInput(urlInput);
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
