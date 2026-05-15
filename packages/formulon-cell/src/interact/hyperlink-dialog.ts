import { type History, recordFormatChange } from '../commands/history.js';
import { clearHyperlink, hyperlinkAt, setHyperlink } from '../commands/hyperlinks.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';

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
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-hldlg__panel');
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

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

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'fc-fmtdlg__btn';
  removeBtn.textContent = t.remove;
  // Anchored on the left side so it doesn't get confused with OK/Cancel.
  removeBtn.style.marginRight = 'auto';
  footer.appendChild(removeBtn);

  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-fmtdlg__btn';
  cancelBtn.textContent = t.cancel;
  footer.appendChild(cancelBtn);

  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
  okBtn.textContent = t.ok;
  footer.appendChild(okBtn);

  const showError = (msg: string): void => {
    errorRow.textContent = msg;
    errorRow.hidden = false;
  };
  const clearError = (): void => {
    errorRow.hidden = true;
    errorRow.textContent = '';
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
      urlInput.focus();
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
        urlInput.focus();
        urlInput.select();
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
