import {
  addThreadedComment,
  commentAt,
  editThreadedComment,
  type listThreadedComments,
  recordCommentChange,
  recordThreadedChange,
  removeThreadedComment,
  resolveThread,
  setComment,
  threadedCommentsAt,
} from '../commands/comment.js';
import type { History } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { cellRect, layoutForView } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { appendDialogActions, appendDialogButton, appendDialogIconButton } from './dialog-shell.js';

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
  open(options?: { mode?: 'note' | 'thread' }): void;
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
  const removeBtn = appendDialogIconButton(head, {
    label: t.remove,
    ariaLabel: t.remove,
    title: t.remove,
    baseClass: 'fc-cmtnote__icon',
    html:
      '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true">' +
      '<path d="M5 2.5a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 .5.5V3h2.5a.5.5 0 0 1 0 1H13v9.5A1.5 1.5 0 0 1 11.5 15h-7A1.5 1.5 0 0 1 3 13.5V4H2.5a.5.5 0 0 1 0-1H5v-.5zM4 4v9.5a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 .5-.5V4H4zm2.5 2a.5.5 0 0 1 .5.5v5a.5.5 0 0 1-1 0v-5a.5.5 0 0 1 .5-.5zm3 0a.5.5 0 0 1 .5.5v5a.5.5 0 0 1-1 0v-5a.5.5 0 0 1 .5-.5z" fill="currentColor"/>' +
      '</svg>',
  });
  wrap.appendChild(head);

  const textarea = document.createElement('textarea');
  textarea.className = 'fc-cmtnote__textarea';
  textarea.placeholder = t.placeholder;
  textarea.setAttribute('aria-label', t.placeholder);
  textarea.spellcheck = true;
  wrap.appendChild(textarea);

  const thread = document.createElement('div');
  thread.className = 'fc-cmtnote__thread';
  thread.hidden = true;
  wrap.insertBefore(thread, textarea);
  const error = document.createElement('p');
  error.className = 'fc-cmtnote__error';
  error.hidden = true;
  wrap.insertBefore(error, textarea);

  const footer = document.createElement('div');
  footer.className = 'fc-cmtnote__footer';
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
    buttonBaseClass: 'fc-cmtnote__btn',
  });
  const replyBtn = appendDialogButton(footer, {
    label: '',
    baseClass: 'fc-cmtnote__btn fc-cmtnote__reply',
  });
  const resolveBtn = appendDialogButton(footer, {
    label: '',
    baseClass: 'fc-cmtnote__btn fc-cmtnote__resolve',
  });
  wrap.appendChild(footer);

  host.appendChild(wrap);

  let threadedMode = false;
  let editingThreadId: string | null = null;
  let replyingTo: string | null = null;
  const labels = {
    threadTitle: t.threadTitle ?? 'Comments',
    threadPlaceholder: t.threadPlaceholder ?? 'Enter a comment',
    reply: t.reply ?? 'Reply',
    resolve: t.resolve ?? 'Resolve',
    reopen: t.reopen ?? 'Reopen',
    edit: t.edit ?? 'Edit',
    noReplies: t.noReplies ?? 'No replies',
    authorFallback: t.authorFallback ?? 'Unknown author',
    noteConflict:
      t.noteConflict ?? 'This cell has a note. Remove it before adding a threaded comment.',
    threadConflict:
      t.threadConflict ?? 'This cell has a threaded comment. Remove it before adding a note.',
    writeFailed: t.writeFailed ?? 'The comment could not be saved.',
  };
  const mutationAuthorization = { origin: 'instanceApi' as const, commandId: 'commentDialog' };
  replyBtn.textContent = labels.reply;
  resolveBtn.textContent = labels.resolve;

  const renderThread = (entries: ReturnType<typeof listThreadedComments>): void => {
    thread.replaceChildren();
    const opening = entries.find((entry) => entry.parentId === '');
    if (!opening) {
      thread.hidden = true;
      return;
    }
    thread.hidden = false;
    for (const entry of entries) {
      const item = document.createElement('article');
      item.className =
        entry.parentId === '' ? 'fc-cmtnote__thread-opening' : 'fc-cmtnote__thread-reply';
      const body = document.createElement('p');
      body.className = 'fc-cmtnote__thread-text';
      body.textContent = entry.text;
      item.appendChild(body);
      const meta = document.createElement('small');
      meta.className = 'fc-cmtnote__thread-meta';
      meta.textContent =
        store.getState().threadedComments.persons.find((person) => person.id === entry.personId)
          ?.displayName ?? labels.authorFallback;
      item.appendChild(meta);
      if (entry.id !== editingThreadId) {
        const edit = appendDialogButton(item, {
          label: labels.edit,
          baseClass: 'fc-cmtnote__thread-edit',
        });
        edit.addEventListener('click', () => {
          editingThreadId = entry.id;
          replyingTo = null;
          textarea.value = entry.text;
          renderThread(entries);
          textarea.focus();
          textarea.select();
        });
      }
      thread.appendChild(item);
    }
    if (entries.length === 1) {
      const empty = document.createElement('small');
      empty.className = 'fc-cmtnote__thread-empty';
      empty.textContent = labels.noReplies;
      thread.appendChild(empty);
    }
    textarea.value = editingThreadId
      ? (entries.find((entry) => entry.id === editingThreadId)?.text ?? '')
      : replyingTo
        ? ''
        : opening.text;
    resolveBtn.textContent = opening.done ? labels.reopen : labels.resolve;
    resolveBtn.hidden = false;
    replyBtn.hidden = false;
    title.textContent = labels.threadTitle;
    textarea.placeholder = labels.threadPlaceholder;
  };

  const writeComment = (next: string | undefined): boolean => {
    const state = store.getState();
    const addr = state.selection.active;
    const wb = getWb() ?? undefined;
    if (
      wb?.capabilities.threadedComments &&
      threadedCommentsAt(store, addr, wb).length > 0 &&
      (next ?? '').length > 0
    ) {
      error.textContent = labels.threadConflict;
      error.hidden = false;
      return false;
    }
    let saved = true;
    recordCommentChange(
      history,
      store,
      wb,
      [addr],
      () => {
        saved = setComment(
          store,
          addr,
          next != null && next !== '' ? next : '',
          wb,
          mutationAuthorization,
        );
      },
      mutationAuthorization,
    );
    if (!saved) {
      error.textContent = labels.writeFailed;
      error.hidden = false;
      return false;
    }
    error.hidden = true;
    return true;
  };

  const writeThread = (): boolean => {
    const state = store.getState();
    const addr = state.selection.active;
    const wb = getWb() ?? undefined;
    if (!wb?.capabilities.threadedComments || textarea.value.length === 0) {
      error.textContent = labels.writeFailed;
      error.hidden = false;
      return false;
    }
    const hasNote = commentAt(state, addr) !== null;
    let result: ReturnType<typeof addThreadedComment> | boolean;
    try {
      result = recordThreadedChange(
        history,
        store,
        wb,
        addr.sheet,
        () => {
          if (editingThreadId) {
            return editThreadedComment(
              store,
              addr.sheet,
              editingThreadId,
              textarea.value,
              wb,
              [],
              mutationAuthorization,
            );
          }
          return addThreadedComment(store, addr, textarea.value, wb, {
            replyToThread: replyingTo ?? undefined,
            authorization: mutationAuthorization,
          });
        },
        mutationAuthorization,
      );
    } catch {
      result = false;
    }
    if (result === null || result === false) {
      error.textContent = hasNote && !editingThreadId ? labels.noteConflict : labels.writeFailed;
      error.hidden = false;
      return false;
    }
    error.hidden = true;
    editingThreadId = null;
    replyingTo = null;
    return true;
  };

  const onOk = (): void => {
    if (threadedMode) {
      if (!writeThread()) return;
    } else if (!writeComment(textarea.value)) return;
    api.close();
  };

  const onRemove = (): void => {
    if (threadedMode) {
      const state = store.getState();
      const addr = state.selection.active;
      const wb = getWb() ?? undefined;
      const opening = threadedCommentsAt(store, addr, wb)[0];
      if (opening && wb) {
        const removed = recordThreadedChange(
          history,
          store,
          wb,
          addr.sheet,
          () => removeThreadedComment(store, addr.sheet, opening.id, wb, mutationAuthorization),
          mutationAuthorization,
        );
        if (!removed) {
          error.textContent = labels.writeFailed;
          error.hidden = false;
          return;
        }
      } else {
        error.textContent = labels.writeFailed;
        error.hidden = false;
        return;
      }
    } else if (!writeComment(undefined)) return;
    api.close();
  };

  const onReply = (): void => {
    if (!threadedMode) return;
    const state = store.getState();
    const addr = state.selection.active;
    const opening = threadedCommentsAt(store, addr, getWb() ?? undefined)[0];
    if (!opening) return;
    editingThreadId = null;
    replyingTo = opening.id;
    textarea.value = '';
    textarea.focus();
  };

  const onResolve = (): void => {
    if (!threadedMode) return;
    const state = store.getState();
    const addr = state.selection.active;
    const wb = getWb() ?? undefined;
    const opening = threadedCommentsAt(store, addr, wb)[0];
    if (opening && wb) {
      const resolved = recordThreadedChange(
        history,
        store,
        wb,
        addr.sheet,
        () =>
          resolveThread(store, addr.sheet, opening.id, !opening.done, wb, mutationAuthorization),
        mutationAuthorization,
      );
      if (!resolved) {
        error.textContent = labels.writeFailed;
        error.hidden = false;
        return;
      }
    } else {
      error.textContent = labels.writeFailed;
      error.hidden = false;
      return;
    }
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
      const r = cellRect(layoutForView(s), s.viewport, a.row, a.col);
      const gridEl = findGrid();
      const offX = gridEl ? gridEl.offsetLeft : 0;
      const offY = gridEl ? gridEl.offsetTop : 0;
      // Cell coords inside the host coordinate system.
      const cellLeft = offX + r.x;
      const cellTop = offY + r.y;
      const cellMid = cellTop + r.h / 2;
      // Prefer the side the sheet reads towards, and fall back to the other
      // when the note would run off the host.
      const rtl = s.ui.rightToLeft === true;
      const preferred = rtl ? cellLeft - GAP - NOTE_W : cellLeft + r.w + GAP;
      const fallback = rtl ? cellLeft + r.w + GAP : cellLeft - GAP - NOTE_W;
      const preferredFits = rtl ? preferred >= 8 : preferred + NOTE_W <= hostW - 8;
      const leftCandidate = preferredFits ? preferred : fallback;
      const placedRight = rtl ? !preferredFits : preferredFits;
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
  replyBtn.addEventListener('click', onReply);
  resolveBtn.addEventListener('click', onResolve);
  cancelBtn.addEventListener('click', onCancel);
  wrap.addEventListener('keydown', onKey);

  let attached = false;
  const onScrollOrResize = (): void => place();

  const api: CommentDialogHandle = {
    open(options: { mode?: 'note' | 'thread' } = {}): void {
      const state = store.getState();
      const addr = state.selection.active;
      const wb = getWb() ?? undefined;
      const current = commentAt(state, addr) ?? '';
      const threads = wb?.capabilities.threadedComments ? threadedCommentsAt(store, addr, wb) : [];
      threadedMode = Boolean(
        wb?.capabilities.threadedComments &&
          options.mode !== 'note' &&
          (options.mode === 'thread' || threads.length > 0 || !current),
      );
      editingThreadId = null;
      replyingTo = null;
      removeBtn.hidden = threadedMode ? threads.length === 0 : !current;
      replyBtn.hidden = !threadedMode || threads.length === 0;
      resolveBtn.hidden = !threadedMode || threads.length === 0;
      title.textContent = threadedMode ? labels.threadTitle : current ? t.titleEdit : t.title;
      wrap.setAttribute('aria-label', title.textContent ?? t.title);
      textarea.placeholder = threadedMode ? labels.threadPlaceholder : t.placeholder;
      textarea.setAttribute('aria-label', threadedMode ? labels.threadPlaceholder : t.placeholder);
      error.hidden = true;
      if (threadedMode && threads.length > 0) renderThread(threads);
      else {
        thread.hidden = true;
        thread.replaceChildren();
        textarea.value = current;
      }
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
      replyBtn.removeEventListener('click', onReply);
      resolveBtn.removeEventListener('click', onResolve);
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
