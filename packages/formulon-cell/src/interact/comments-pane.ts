import { listThreadedComments } from '../commands/comment.js';
import { formatA1Cell } from '../engine/address.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { MAC_CONTEXT_LABELS } from '../toolbar/ribbon/mac/labels-data.js';
import { toolbarLangForLocale } from '../toolbar/ribbon/mac/locale.js';
import { appendDialogButton } from './dialog-shell.js';
import { isNavigationAddrAllowed } from './navigation-policy.js';
import { syncTaskpaneDock } from './taskpane-dock.js';

export interface CommentsPaneDeps {
  /** The `.fc-host` element; the pane docks in its task-pane dock. */
  host: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle;
  getStrings: () => Strings;
  getLocale: () => string;
}

export interface CommentsPaneHandle {
  detach(): void;
}

/**
 * The Mac Comments task pane: the active sheet's threads, each under its cell, leaving out
 * cells the navigation policy keeps out of reach.
 * Its visibility is `ui.commentsPaneOpen`, so every opener and closer shares one state.
 */
export function attachCommentsPane(deps: CommentsPaneDeps): CommentsPaneHandle {
  const { host, store, getWb } = deps;
  const dock = host.querySelector<HTMLElement>('.fc-host__taskpane-dock');

  const root = document.createElement('div');
  root.className = 'fc-commentspane';
  root.setAttribute('role', 'region');
  root.hidden = !store.getState().ui.commentsPaneOpen;

  const header = document.createElement('div');
  header.className = 'fc-commentspane__header';
  const title = document.createElement('span');
  title.className = 'fc-commentspane__title';
  const closeBtn = appendDialogButton(header, { label: '', baseClass: 'fc-commentspane__close' });
  closeBtn.append(createExcelRibbonSvg('paneClose', 'fc-pane-close__glyph') ?? '');
  header.prepend(title);
  const body = document.createElement('div');
  body.className = 'fc-commentspane__body';
  root.append(header, body);
  (dock ?? host).appendChild(root);

  let lastOpen = root.hidden === false;
  let lastSheet = store.getState().data.sheetIndex;
  let lastThreads = store.getState().threadedComments.bySheet;
  let rendering = false;
  const render = (): void => {
    rendering = true;
    try {
      draw();
    } finally {
      rendering = false;
      lastThreads = store.getState().threadedComments.bySheet;
    }
  };
  const draw = (): void => {
    const heading = MAC_CONTEXT_LABELS.commentsTitle[toolbarLangForLocale(deps.getLocale())];
    title.textContent = heading;
    root.setAttribute('aria-label', heading);
    closeBtn.setAttribute('aria-label', deps.getStrings().workbookObjects.close);
    const sheet = store.getState().data.sheetIndex;
    const entries = listThreadedComments(store, sheet, getWb()).filter((entry) =>
      isNavigationAddrAllowed(store, { sheet, row: entry.cell.row, col: entry.cell.col }),
    );
    body.replaceChildren();
    const roots = entries.filter((entry) => !entry.parentId);
    if (roots.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'fc-commentspane__empty';
      empty.textContent = deps.getStrings().mac.commentsPane.empty;
      body.appendChild(empty);
      return;
    }
    for (const entry of roots) {
      const item = document.createElement('div');
      item.className = 'fc-commentspane__item';
      const cell = document.createElement('strong');
      cell.className = 'fc-commentspane__cell';
      cell.textContent = formatA1Cell(entry.cell.row, entry.cell.col);
      item.appendChild(cell);
      for (const part of [entry, ...entries.filter((reply) => reply.parentId === entry.id)]) {
        const text = document.createElement('p');
        text.className = 'fc-commentspane__text';
        text.textContent = part.text;
        item.appendChild(text);
      }
      body.appendChild(item);
    }
  };

  const sync = (): void => {
    if (rendering) return;
    const s = store.getState();
    const open = s.ui.commentsPaneOpen;
    const changed =
      open !== lastOpen ||
      s.data.sheetIndex !== lastSheet ||
      s.threadedComments.bySheet !== lastThreads;
    lastSheet = s.data.sheetIndex;
    lastThreads = s.threadedComments.bySheet;
    if (!changed) return;
    if (open !== lastOpen) {
      lastOpen = open;
      root.hidden = !open;
      if (dock) syncTaskpaneDock(dock);
    }
    if (open) render();
  };
  const unsubscribe = store.subscribe(sync);
  closeBtn.addEventListener('click', () => {
    mutators.setCommentsPaneOpen(store, false);
    host.focus();
  });
  if (!root.hidden) render();
  if (dock) syncTaskpaneDock(dock);

  return {
    detach: () => {
      unsubscribe();
      root.remove();
      if (dock) syncTaskpaneDock(dock);
    },
  };
}
