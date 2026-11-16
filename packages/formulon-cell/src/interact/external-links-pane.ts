import { type ExternalLinkKind, listExternalLinks } from '../commands/external-links.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { focusSheet, GRID_SELECTOR } from './dialog-focus.js';
import { bindPaneEscape, markSurface } from './dialog-shell.js';
import { createDockedPane, syncTaskpaneDock, taskpaneDockOf } from './taskpane-dock.js';

/** Wording of the docked Workbook Links pane. */
export interface ExternalLinksPaneStrings {
  title: string;
  close: string;
  empty: string;
  help: string;
}

export interface ExternalLinksPaneDeps {
  host: HTMLElement;
  getWb: () => WorkbookHandle | null;
  strings?: Strings;
  paneStrings: () => ExternalLinksPaneStrings;
}

export interface ExternalLinksPaneHandle {
  open(): void;
  close(): void;
  refresh(): void;
  detach(): void;
}

const kindText = (kind: ExternalLinkKind, labels: Strings['externalLinksDialog']): string => {
  switch (kind) {
    case 'externalBook':
      return labels.kindExternalBook;
    case 'ole':
      return labels.kindOle;
    case 'dde':
      return labels.kindDde;
    case 'unknown':
      return labels.kindUnknown;
  }
};

/**
 * Workbook Links as a pane docked to the right edge of the grid: a header tab
 * with a round close button, the workbook's link list or its empty-state
 * sentence, and a help link.
 */
export function attachExternalLinksPane(deps: ExternalLinksPaneDeps): ExternalLinksPaneHandle {
  const { host, getWb } = deps;
  let strings = deps.strings ?? defaultStrings;
  let wording = deps.paneStrings();

  const { root, title, closeBtn, body } = createDockedPane('fc-extlinkpane');
  root.setAttribute('aria-label', wording.title);
  root.hidden = true;

  const list = document.createElement('div');
  list.className = 'fc-extlinkpane__list';
  const empty = document.createElement('p');
  empty.className = 'fc-extlinkpane__empty';
  const help = document.createElement('span');
  help.className = 'fc-extlinkpane__help';
  help.setAttribute('role', 'link');
  help.tabIndex = 0;
  body.append(empty, list, help);
  host.appendChild(root);
  markSurface(root, 'pane');
  const detachPaneEscape = bindPaneEscape(root, host);

  const labels = (): Strings['externalLinksDialog'] => strings.externalLinksDialog;

  // The pane takes the host's task-pane dock beside the grid, as the other
  // docked panes do; a host without one gets it laid over the grid's right edge.
  const dock = (): void => {
    const target = taskpaneDockOf(host);
    if (target) {
      if (root.parentElement !== target) target.appendChild(root);
      root.style.removeProperty('top');
      root.style.removeProperty('bottom');
      return;
    }
    const grid = host.querySelector<HTMLElement>(GRID_SELECTOR);
    root.style.top = `${grid?.offsetTop ?? 0}px`;
    root.style.bottom = grid
      ? `${host.clientHeight - grid.offsetTop - grid.offsetHeight}px`
      : '0px';
  };

  const render = (): void => {
    title.textContent = wording.title;
    closeBtn.setAttribute('aria-label', wording.close);
    closeBtn.title = wording.close;
    empty.textContent = wording.empty;
    help.textContent = wording.help;
    root.setAttribute('aria-label', wording.title);
    const links = listExternalLinks(getWb());
    empty.hidden = links.length > 0;
    list.replaceChildren(
      ...links.map((link) => {
        const row = document.createElement('div');
        row.className = 'fc-extlinkpane__link';
        const target = document.createElement('span');
        target.className = 'fc-extlinkpane__target';
        target.textContent = link.target || '—';
        target.title = link.target;
        const kind = document.createElement('span');
        kind.className = 'fc-extlinkpane__kind';
        kind.textContent = kindText(link.kind, labels());
        row.append(target, kind);
        return row;
      }),
    );
  };

  const hide = (): void => {
    root.hidden = true;
    const target = taskpaneDockOf(host);
    if (target) syncTaskpaneDock(target);
  };

  closeBtn.addEventListener('click', () => {
    hide();
    focusSheet(host);
  });

  return {
    open() {
      render();
      dock();
      root.hidden = false;
      const target = taskpaneDockOf(host);
      if (target) syncTaskpaneDock(target);
    },
    close() {
      hide();
    },
    refresh() {
      strings = deps.strings ?? defaultStrings;
      wording = deps.paneStrings();
      render();
    },
    detach() {
      detachPaneEscape();
      root.remove();
      const target = taskpaneDockOf(host);
      if (target) syncTaskpaneDock(target);
    },
  };
}
