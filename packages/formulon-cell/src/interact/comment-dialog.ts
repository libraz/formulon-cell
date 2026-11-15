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
import { commentWrite, threadWrite } from '../commands/comment-actions.js';
import { DEFAULT_COMMENT_AUTHOR } from '../commands/comment-author.js';
import type { History } from '../commands/history.js';
import { formatA1Cell, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { cellRect, layoutForView, pointPx } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { type DialogEntry, INSTANCE_API_COMMENT_ENTRY } from './dialog-entry.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogIconButton,
  deferInitialFocus,
  markSurface,
  openerOf,
  returnFocus,
} from './dialog-shell.js';
import { onOutsidePress } from './outside-press.js';

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
  open(options?: { mode?: 'note' | 'thread'; entry?: DialogEntry }): void;
  close(): void;
  detach(): void;
}

const NOTE_W = 240;
const NOTE_MIN_H = 132;
// The Mac conversation card, in pt taken as px; it centres on the cell.
const START_W = 244;
const START_H = 104;
const GAP = 8;
// The Mac note shape in pt, scaled to screen px by `pointPx`: size, gap right of the cell, lift above
// the cell top, and the distance from the sheet bottom when flipped up.
const MAC_NOTE_W = 104;
const MAC_NOTE_H = 76;
const MAC_NOTE_GAP = 15;
const MAC_NOTE_LIFT = 10;
const MAC_NOTE_FLIP_UP = 88;
const MAC_NOTE_FONT_PT = 10;
// Trims float noise from pt-to-px products.
const px = (v: number): number => Math.round(v * 100) / 100;

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
  markSurface(wrap, 'floating');
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

  // Mac: a new conversation opens as a card headed by author, cell and close.
  const start = document.createElement('div');
  start.className = 'fc-cmtnote__start';
  start.hidden = true;
  const avatar = document.createElement('span');
  avatar.className = 'fc-cmtnote__start-avatar';
  avatar.setAttribute('aria-hidden', 'true');
  const avatarGlyph = createExcelRibbonSvg('commentAvatar', 'fc-cmtnote__start-glyph');
  if (avatarGlyph) avatar.appendChild(avatarGlyph);
  const startName = document.createElement('strong');
  startName.className = 'fc-cmtnote__start-name';
  const startAddr = document.createElement('span');
  startAddr.className = 'fc-cmtnote__start-addr';
  start.append(avatar, startName, startAddr);
  const startClose = appendDialogIconButton(start, {
    label: '',
    ariaLabel: t.close,
    title: t.close,
    baseClass: 'fc-cmtnote__start-close',
  });
  wrap.appendChild(start);

  const textarea = document.createElement('textarea');
  textarea.className = 'fc-cmtnote__textarea';
  textarea.placeholder = t.placeholder;
  textarea.setAttribute('aria-label', t.placeholder);
  textarea.spellcheck = true;
  wrap.appendChild(textarea);
  const sendBtn = appendDialogIconButton(wrap, {
    label: '',
    ariaLabel: t.send,
    title: t.send,
    baseClass: 'fc-cmtnote__send',
  });
  sendBtn.hidden = true;
  const sendGlyph = createExcelRibbonSvg('commentSend', 'fc-cmtnote__start-glyph');
  if (sendGlyph) sendBtn.appendChild(sendGlyph);

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
  let opener: HTMLElement | null = null;

  let threadedMode = false;
  let startMode = false;
  let macNote = false;
  // The cell the dialog opened on: a click that commits the note may already have moved the selection.
  let target = store.getState().selection.active;
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
  // The entry point that opened the dialog; its commits run as it.
  let mutationAuthorization: DialogEntry = INSTANCE_API_COMMENT_ENTRY;
  const commentContext = () => ({ store, history: history ?? null, ...mutationAuthorization });
  replyBtn.textContent = labels.reply;
  resolveBtn.textContent = labels.resolve;

  const renderThread = (entries: ReturnType<typeof listThreadedComments>): void => {
    thread.replaceChildren();
    const opening = entries.find((entry) => entry.parentId === null);
    if (!opening) {
      thread.hidden = true;
      return;
    }
    thread.hidden = false;
    for (const entry of entries) {
      const item = document.createElement('article');
      item.className =
        entry.parentId === null ? 'fc-cmtnote__thread-opening' : 'fc-cmtnote__thread-reply';
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
    const addr = target;
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
    let saved = false;
    commentWrite(commentContext(), [addr], undefined, () =>
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
      ),
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
    const addr = target;
    const wb = getWb() ?? undefined;
    if (!wb?.capabilities.threadedComments || textarea.value.length === 0) {
      error.textContent = labels.writeFailed;
      error.hidden = false;
      return false;
    }
    const hasNote = commentAt(state, addr) !== null;
    let result: ReturnType<typeof addThreadedComment> | boolean;
    try {
      result = threadWrite(commentContext(), [addr], false, () =>
        recordThreadedChange(
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
        ),
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
      const addr = target;
      const wb = getWb() ?? undefined;
      const opening = threadedCommentsAt(store, addr, wb)[0];
      if (opening && wb) {
        const removed = threadWrite(commentContext(), [addr], false, () =>
          recordThreadedChange(
            history,
            store,
            wb,
            addr.sheet,
            () => removeThreadedComment(store, addr.sheet, opening.id, wb, mutationAuthorization),
            mutationAuthorization,
          ),
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
    const addr = target;
    const opening = threadedCommentsAt(store, addr, getWb() ?? undefined)[0];
    if (!opening) return;
    editingThreadId = null;
    replyingTo = opening.id;
    textarea.value = '';
    textarea.focus();
  };

  const onResolve = (): void => {
    if (!threadedMode) return;
    const addr = target;
    const wb = getWb() ?? undefined;
    const opening = threadedCommentsAt(store, addr, wb)[0];
    if (opening && wb) {
      const resolved = threadWrite(commentContext(), [addr], false, () =>
        recordThreadedChange(
          history,
          store,
          wb,
          addr.sheet,
          () =>
            resolveThread(store, addr.sheet, opening.id, !opening.done, wb, mutationAuthorization),
          mutationAuthorization,
        ),
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
    const a = target;
    const hostW = host.clientWidth;
    const z = macNote ? pointPx(s.layout, s.viewport) : 1;
    const noteW = startMode ? START_W : macNote ? MAC_NOTE_W * z : NOTE_W;
    const gap = macNote ? MAC_NOTE_GAP * z : GAP;
    let left = Math.max(8, hostW / 2 - noteW / 2);
    let top = 80;
    let tailLeft: number | null = null;
    try {
      const r = cellRect(layoutForView(s), s.viewport, a.row, a.col);
      // The grid's offsetParent can be a wrapper inside the host, so measure against the host.
      const gridEl = findGrid();
      const gridBox = gridEl?.getBoundingClientRect();
      const hostBox = host.getBoundingClientRect();
      const offX = gridBox ? gridBox.left - hostBox.left - host.clientLeft : 0;
      const offY = gridBox ? gridBox.top - hostBox.top - host.clientTop : 0;
      // Cell coords inside the host coordinate system.
      const cellLeft = offX + r.x;
      const cellTop = offY + r.y;
      const cellMid = cellTop + r.h / 2;
      // Prefer the side the sheet reads towards, and fall back to the other
      // when the note would run off the host.
      const rtl = s.ui.rightToLeft === true;
      const preferred = rtl ? cellLeft - gap - noteW : cellLeft + r.w + gap;
      const fallback = rtl ? cellLeft + r.w + gap : cellLeft - gap - noteW;
      const sheetEdge = macNote && !rtl && a.col >= MAX_COL;
      const preferredFits = !sheetEdge && (rtl ? preferred >= 8 : preferred + noteW <= hostW - 8);
      const leftCandidate = preferredFits ? preferred : fallback;
      const placedRight = rtl ? !preferredFits : preferredFits;
      left = Math.max(8, leftCandidate);
      // Keep the note vertically near the cell — the tail is positioned at
      // top:14 inside the note, so we line that up with the cell mid.
      if (macNote) {
        const flipTop = a.row >= MAX_ROW ? cellTop + r.h - MAC_NOTE_FLIP_UP * z : null;
        top = Math.max(8, flipTop ?? cellTop - MAC_NOTE_LIFT * z);
      } else top = Math.max(8, startMode ? cellMid - START_H / 2 : cellMid - 14 - 6);
      tailLeft = placedRight ? -7 : noteW - 1;
    } catch {
      tailLeft = null;
    }
    wrap.style.left = `${Math.round(left)}px`;
    wrap.style.top = `${Math.round(top)}px`;
    wrap.style.width = `${px(noteW)}px`;
    if (macNote) {
      wrap.style.minHeight = '';
      wrap.style.height = `${px(MAC_NOTE_H * z)}px`;
      wrap.style.fontSize = `${px(MAC_NOTE_FONT_PT * z)}px`;
    } else {
      wrap.style.height = '';
      wrap.style.fontSize = '';
      wrap.style.minHeight = `${startMode ? START_H : NOTE_MIN_H}px`;
    }
    if (tailLeft == null) {
      tail.style.display = 'none';
    } else {
      tail.style.display = '';
      tail.style.left = `${tailLeft}px`;
      tail.classList.toggle('fc-cmtnote__tail--right', tailLeft > 0);
    }
  };

  okBtn.addEventListener('click', onOk);
  sendBtn.addEventListener('click', onOk);
  startClose.addEventListener('click', onCancel);
  textarea.addEventListener('input', () => {
    if (startMode) projectDisabledState(sendBtn, textarea.value.length === 0, null);
  });
  removeBtn.addEventListener('click', onRemove);
  replyBtn.addEventListener('click', onReply);
  resolveBtn.addEventListener('click', onResolve);
  cancelBtn.addEventListener('click', onCancel);
  wrap.addEventListener('keydown', onKey);

  let attached = false;
  let detachOutside: (() => void) | null = null;
  let placeFrame = 0;
  const cancelPlaceFrame = (): void => {
    if (placeFrame) window.cancelAnimationFrame(placeFrame);
    placeFrame = 0;
  };
  // Coalesce to one reposition per frame; scrolling inside the panel does not move it.
  const onScrollOrResize = (e: Event): void => {
    if (e.type === 'scroll' && e.target instanceof Node && wrap.contains(e.target)) return;
    if (placeFrame) return;
    placeFrame = window.requestAnimationFrame(() => {
      placeFrame = 0;
      place();
    });
  };

  const api: CommentDialogHandle = {
    open(options: { mode?: 'note' | 'thread'; entry?: DialogEntry } = {}): void {
      mutationAuthorization = options.entry ?? INSTANCE_API_COMMENT_ENTRY;
      if (!wrap.contains(document.activeElement)) opener = openerOf(document);
      const state = store.getState();
      const addr = state.selection.active;
      target = addr;
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
      const macHost =
        (host.closest?.<HTMLElement>('.fc-host') ?? host).dataset?.fcPlatform === 'mac';
      startMode = threadedMode && threads.length === 0 && macHost;
      // A new note on the Mac profile is a bare text box with no header or footer.
      macNote = !startMode && !threadedMode && !current && macHost;
      wrap.classList.toggle('fc-cmtnote--start', startMode);
      wrap.classList.toggle('fc-cmtnote--mac-note', macNote);
      head.hidden = startMode || macNote;
      footer.hidden = startMode || macNote;
      start.hidden = !startMode;
      sendBtn.hidden = !startMode;
      projectDisabledState(sendBtn, true, null);
      startName.textContent = DEFAULT_COMMENT_AUTHOR;
      startAddr.textContent = formatA1Cell(addr.row, addr.col);
      removeBtn.hidden = threadedMode ? threads.length === 0 : !current;
      replyBtn.hidden = !threadedMode || threads.length === 0;
      resolveBtn.hidden = !threadedMode || threads.length === 0;
      title.textContent = threadedMode ? labels.threadTitle : current ? t.titleEdit : t.title;
      wrap.setAttribute('aria-label', title.textContent ?? t.title);
      const placeholder = startMode
        ? t.threadStart
        : threadedMode
          ? labels.threadPlaceholder
          : t.placeholder;
      textarea.placeholder = placeholder;
      textarea.setAttribute('aria-label', placeholder);
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
        // An outside press cancels the note; a Mac new note commits instead.
        detachOutside = onOutsidePress(
          document,
          () => [wrap],
          () => {
            if (wrap.hidden) return;
            if (macNote) onOk();
            else onCancel();
          },
          { event: 'mousedown' },
        );
        window.addEventListener('scroll', onScrollOrResize, true);
        window.addEventListener('resize', onScrollOrResize);
        attached = true;
      }
      deferInitialFocus(wrap, () => {
        textarea.focus();
        textarea.select();
      });
    },
    close(): void {
      wrap.hidden = true;
      if (attached) {
        detachOutside?.();
        detachOutside = null;
        window.removeEventListener('scroll', onScrollOrResize, true);
        window.removeEventListener('resize', onScrollOrResize);
        attached = false;
      }
      cancelPlaceFrame();
      const target = opener;
      opener = null;
      returnFocus(target, wrap, host);
    },
    detach(): void {
      okBtn.removeEventListener('click', onOk);
      removeBtn.removeEventListener('click', onRemove);
      replyBtn.removeEventListener('click', onReply);
      resolveBtn.removeEventListener('click', onResolve);
      cancelBtn.removeEventListener('click', onCancel);
      wrap.removeEventListener('keydown', onKey);
      if (attached) {
        detachOutside?.();
        detachOutside = null;
        window.removeEventListener('scroll', onScrollOrResize, true);
        window.removeEventListener('resize', onScrollOrResize);
        attached = false;
      }
      cancelPlaceFrame();
      wrap.remove();
    },
  };

  return api;
}
