import { clearComment, commentAt, setComment } from '../commands/comment.js';
import { type History, recordFormatChange } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';

export interface CommentDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  history?: History | null;
  /** Workbook getter — lazy so the dialog stays in lockstep with `setWorkbook`
   *  swaps. The OK/Remove path flushes the change to the engine for xlsx
   *  round-trip when supported. */
  getWb?: () => WorkbookHandle | null;
}

export interface CommentDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

/**
 * Edit-comment dialog. Excel 365 styled callout that replaces native
 * `window.prompt` for setting/clearing the per-cell comment on the active
 * cell. Multi-cell apply still goes through the full Format Cells dialog.
 */
export function attachCommentDialog(deps: CommentDialogDeps): CommentDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  const strings = deps.strings ?? defaultStrings;
  const t = strings.commentDialog;

  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg fc-cmtdlg';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel fc-cmtdlg__panel';
  overlay.appendChild(panel);

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const label = document.createElement('label');
  label.textContent = t.placeholder;
  label.className = 'fc-cmtdlg__label';
  const textarea = document.createElement('textarea');
  textarea.className = 'fc-fmtdlg__textarea fc-cmtdlg__textarea';
  textarea.rows = 5;
  textarea.placeholder = t.placeholder;
  textarea.spellcheck = true;
  label.appendChild(textarea);
  row.appendChild(label);
  body.appendChild(row);

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'fc-fmtdlg__btn';
  removeBtn.textContent = t.remove;
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

  const writeComment = (next: string | undefined): void => {
    const state = store.getState();
    const addr = state.selection.active;
    const wb = getWb() ?? undefined;
    recordFormatChange(history, store, () => {
      if (next != null && next !== '') setComment(store, addr, next, wb);
      else clearComment(store, addr, wb);
    });
  };

  const onOk = (): void => {
    const value = textarea.value;
    writeComment(value);
    api.close();
  };

  const onRemove = (): void => {
    writeComment(undefined);
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
    // Ctrl/Cmd + Enter commits — plain Enter inserts a newline (Excel parity).
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      onOk();
    }
  };

  okBtn.addEventListener('click', onOk);
  removeBtn.addEventListener('click', onRemove);
  cancelBtn.addEventListener('click', onCancel);
  overlay.addEventListener('click', onOverlayClick);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: CommentDialogHandle = {
    open(): void {
      const state = store.getState();
      const addr = state.selection.active;
      const current = commentAt(state, addr) ?? '';
      textarea.value = current;
      removeBtn.hidden = !current;
      header.textContent = current ? t.titleEdit : t.title;
      overlay.setAttribute('aria-label', current ? t.titleEdit : t.title);
      overlay.hidden = false;
      requestAnimationFrame(() => {
        textarea.focus();
        textarea.select();
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
