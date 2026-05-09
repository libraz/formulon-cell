import { clearComment, commentAt, setComment } from '../commands/comment.js';
import { type History, recordFormatChange } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { cellRect } from '../render/geometry.js';
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

const NOTE_W = 240;
const NOTE_MIN_H = 132;
const GAP = 8;

/**
 * Edit-comment popover. Renders as an desktop-spreadsheet style yellow sticky note
 * anchored to the active cell — no centered modal, no dark backdrop. Click
 * outside / Escape commits-as-cancel; Enter inserts a newline; Ctrl/Cmd +
 * Enter or the OK button commits; the trash icon clears the note.
 */
export function attachCommentDialog(deps: CommentDialogDeps): CommentDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const getWb = deps.getWb ?? ((): WorkbookHandle | null => null);
  const strings = deps.strings ?? defaultStrings;
  const t = strings.commentDialog;

  const wrap = document.createElement('div');
  wrap.className = 'fc-cmtnote';
  wrap.setAttribute('role', 'dialog');
  wrap.setAttribute('aria-modal', 'false');
  wrap.setAttribute('aria-label', t.title);
  wrap.hidden = true;

  const tail = document.createElement('span');
  tail.className = 'fc-cmtnote__tail';
  tail.setAttribute('aria-hidden', 'true');
  wrap.appendChild(tail);

  const head = document.createElement('div');
  head.className = 'fc-cmtnote__head';
  const title = document.createElement('span');
  title.className = 'fc-cmtnote__title';
  title.textContent = t.title;
  head.appendChild(title);
  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'fc-cmtnote__icon';
  removeBtn.setAttribute('aria-label', t.remove);
  removeBtn.title = t.remove;
  removeBtn.innerHTML =
    '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true">' +
    '<path d="M5 2.5a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 .5.5V3h2.5a.5.5 0 0 1 0 1H13v9.5A1.5 1.5 0 0 1 11.5 15h-7A1.5 1.5 0 0 1 3 13.5V4H2.5a.5.5 0 0 1 0-1H5v-.5zM4 4v9.5a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 .5-.5V4H4zm2.5 2a.5.5 0 0 1 .5.5v5a.5.5 0 0 1-1 0v-5a.5.5 0 0 1 .5-.5zm3 0a.5.5 0 0 1 .5.5v5a.5.5 0 0 1-1 0v-5a.5.5 0 0 1 .5-.5z" fill="currentColor"/>' +
    '</svg>';
  head.appendChild(removeBtn);
  wrap.appendChild(head);

  const textarea = document.createElement('textarea');
  textarea.className = 'fc-cmtnote__textarea';
  textarea.placeholder = t.placeholder;
  textarea.spellcheck = true;
  wrap.appendChild(textarea);

  const footer = document.createElement('div');
  footer.className = 'fc-cmtnote__footer';
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-cmtnote__btn';
  cancelBtn.textContent = t.cancel;
  footer.appendChild(cancelBtn);
  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-cmtnote__btn fc-cmtnote__btn--primary';
  okBtn.textContent = t.ok;
  footer.appendChild(okBtn);
  wrap.appendChild(footer);

  host.appendChild(wrap);

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
    writeComment(textarea.value);
    api.close();
  };

  const onRemove = (): void => {
    writeComment(undefined);
    api.close();
  };

  const onCancel = (): void => api.close();

  // Click-outside guard — committed cancel when the user clicks anywhere
  // that isn't the note. Captured at document level so we beat host-bound
  // pointer handlers (selection, autofill, etc.).
  const onDocPointerDown = (e: MouseEvent): void => {
    if (wrap.hidden) return;
    if (e.target instanceof Node && wrap.contains(e.target)) return;
    onCancel();
  };

  const onKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
      return;
    }
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      onOk();
    }
  };

  // Locate the grid canvas — cellRect returns coords relative to the canvas,
  // so we need its offset within the host to position correctly.
  const findGrid = (): HTMLElement | null => {
    const el = host.querySelector('.fc-host__grid');
    return el instanceof HTMLElement ? el : null;
  };

  // Place the note next to the active cell. Coordinates are host-relative —
  // the host element acts as the positioning ancestor (it carries the
  // overlay's containing block). Falls back to the host's centre if the
  // geometry isn't yet available (e.g. on first mount before a paint).
  const place = (): void => {
    const s = store.getState();
    const a = s.selection.active;
    const hostW = host.clientWidth;
    let left = Math.max(8, hostW / 2 - NOTE_W / 2);
    let top = 80;
    let tailLeft: number | null = null;
    try {
      const r = cellRect(s.layout, s.viewport, a.row, a.col);
      const gridEl = findGrid();
      const offX = gridEl ? gridEl.offsetLeft : 0;
      const offY = gridEl ? gridEl.offsetTop : 0;
      // Cell coords inside the host coordinate system.
      const cellLeft = offX + r.x;
      const cellTop = offY + r.y;
      const cellMid = cellTop + r.h / 2;
      // Prefer right of the cell; fall back to the left if it would overflow.
      let leftCandidate = cellLeft + r.w + GAP;
      const willOverflowRight = leftCandidate + NOTE_W > hostW - 8;
      if (willOverflowRight) leftCandidate = cellLeft - GAP - NOTE_W;
      const placedRight = !willOverflowRight;
      left = Math.max(8, leftCandidate);
      // Keep the note vertically near the cell — the tail is positioned at
      // top:14 inside the note, so we line that up with the cell mid.
      top = Math.max(8, cellMid - 14 - 6);
      tailLeft = placedRight ? -7 : NOTE_W - 1;
    } catch {
      tailLeft = null;
    }
    wrap.style.left = `${Math.round(left)}px`;
    wrap.style.top = `${Math.round(top)}px`;
    wrap.style.width = `${NOTE_W}px`;
    wrap.style.minHeight = `${NOTE_MIN_H}px`;
    if (tailLeft == null) {
      tail.style.display = 'none';
    } else {
      tail.style.display = '';
      tail.style.left = `${tailLeft}px`;
      tail.classList.toggle('fc-cmtnote__tail--right', tailLeft > 0);
    }
  };

  okBtn.addEventListener('click', onOk);
  removeBtn.addEventListener('click', onRemove);
  cancelBtn.addEventListener('click', onCancel);
  wrap.addEventListener('keydown', onKey);

  let attached = false;
  const onScrollOrResize = (): void => place();

  const api: CommentDialogHandle = {
    open(): void {
      const state = store.getState();
      const addr = state.selection.active;
      const current = commentAt(state, addr) ?? '';
      textarea.value = current;
      removeBtn.hidden = !current;
      title.textContent = current ? t.titleEdit : t.title;
      wrap.setAttribute('aria-label', current ? t.titleEdit : t.title);
      wrap.hidden = false;
      place();
      if (!attached) {
        document.addEventListener('mousedown', onDocPointerDown, true);
        window.addEventListener('scroll', onScrollOrResize, true);
        window.addEventListener('resize', onScrollOrResize);
        attached = true;
      }
      requestAnimationFrame(() => {
        textarea.focus();
        textarea.select();
      });
    },
    close(): void {
      wrap.hidden = true;
      if (attached) {
        document.removeEventListener('mousedown', onDocPointerDown, true);
        window.removeEventListener('scroll', onScrollOrResize, true);
        window.removeEventListener('resize', onScrollOrResize);
        attached = false;
      }
      host.focus();
    },
    detach(): void {
      okBtn.removeEventListener('click', onOk);
      removeBtn.removeEventListener('click', onRemove);
      cancelBtn.removeEventListener('click', onCancel);
      wrap.removeEventListener('keydown', onKey);
      if (attached) {
        document.removeEventListener('mousedown', onDocPointerDown, true);
        window.removeEventListener('scroll', onScrollOrResize, true);
        window.removeEventListener('resize', onScrollOrResize);
        attached = false;
      }
      wrap.remove();
    },
  };

  return api;
}
