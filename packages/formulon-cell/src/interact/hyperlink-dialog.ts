import { type History, recordFormatChange } from '../commands/history.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

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
 * Excel-style "Insert Hyperlink" dialog. Edits the hyperlink field on the
 * currently active cell only — multi-cell apply goes through the full Format
 * Cells dialog instead.
 */
export function attachHyperlinkDialog(deps: HyperlinkDialogDeps): HyperlinkDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  const strings = deps.strings ?? defaultStrings;
  const t = strings.hyperlinkDialog;

  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg fc-hldlg';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel fc-hldlg__panel';
  overlay.appendChild(panel);

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

  host.appendChild(overlay);

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
    const range = { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col };
    recordFormatChange(history, store, () => {
      mutators.setRangeFormat(store, range, { hyperlink: next });
    });
    const wb = getWb();
    if (wb) flushFormatToEngine(wb, store, addr.sheet);
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

  const onOverlayClick = (e: MouseEvent): void => {
    if (e.target === overlay) api.close();
  };

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

  okBtn.addEventListener('click', onOk);
  removeBtn.addEventListener('click', onRemove);
  cancelBtn.addEventListener('click', onCancel);
  overlay.addEventListener('click', onOverlayClick);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: HyperlinkDialogHandle = {
    open(): void {
      const state = store.getState();
      const addr = state.selection.active;
      const fmt = state.format.formats.get(addrKey(addr));
      const current = fmt?.hyperlink ?? '';
      urlInput.value = current;
      // Hide Remove when there's nothing to remove.
      removeBtn.hidden = !current;
      clearError();
      overlay.hidden = false;
      requestAnimationFrame(() => {
        urlInput.focus();
        urlInput.select();
      });
    },
    close(): void {
      overlay.hidden = true;
      host.focus();
    },
    detach(): void {
      okBtn.removeEventListener('click', onOk);
      removeBtn.removeEventListener('click', onRemove);
      cancelBtn.removeEventListener('click', onCancel);
      overlay.removeEventListener('click', onOverlayClick);
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };

  return api;
}
