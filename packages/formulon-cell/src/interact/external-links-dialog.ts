import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';

export interface ExternalLinksDialogDeps {
  host: HTMLElement;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so opening it always reads the live link list. */
  getWb: () => WorkbookHandle | null;
  strings?: Strings;
}

export interface ExternalLinksDialogHandle {
  open(): void;
  close(): void;
  /** Re-read i18n strings (e.g. after a locale switch). */
  refresh(): void;
  detach(): void;
}

/**
 * Excel-parity "Edit Links" — read-only inventory of `<externalReferences>`
 * records carried by the workbook. Records that round-trip through formulon's
 * passthrough mechanism stay listed but are not editable through this UI.
 */
export function attachExternalLinksDialog(
  deps: ExternalLinksDialogDeps,
): ExternalLinksDialogHandle {
  const { host, getWb } = deps;
  let strings = deps.strings ?? defaultStrings;
  let t = strings.externalLinksDialog;

  const overlay = document.createElement('div');
  overlay.className = 'fc-extlinkdlg';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = 'fc-extlinkdlg__panel';
  overlay.appendChild(panel);

  const header = document.createElement('div');
  header.className = 'fc-extlinkdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-extlinkdlg__body';
  panel.appendChild(body);

  const note = document.createElement('p');
  note.className = 'fc-extlinkdlg__note';
  note.textContent = t.note;
  body.appendChild(note);

  const tableWrap = document.createElement('div');
  tableWrap.className = 'fc-extlinkdlg__tablewrap';
  body.appendChild(tableWrap);

  const empty = document.createElement('div');
  empty.className = 'fc-extlinkdlg__empty';
  empty.textContent = t.empty;
  empty.hidden = true;
  body.appendChild(empty);

  const footer = document.createElement('div');
  footer.className = 'fc-extlinkdlg__footer';
  panel.appendChild(footer);

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-extlinkdlg__close';
  closeBtn.textContent = t.close;
  closeBtn.addEventListener('click', () => close());
  footer.appendChild(closeBtn);

  host.appendChild(overlay);

  const onOverlayClick = (e: MouseEvent): void => {
    if (e.target === overlay) close();
  };
  overlay.addEventListener('click', onOverlayClick);

  const onKey = (e: KeyboardEvent): void => {
    if (overlay.hidden) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };
  document.addEventListener('keydown', onKey);

  const renderTable = (): void => {
    tableWrap.replaceChildren();
    const wb = getWb();
    const links = wb?.getExternalLinks() ?? [];
    if (links.length === 0) {
      empty.hidden = false;
      return;
    }
    empty.hidden = true;
    const table = document.createElement('table');
    table.className = 'fc-extlinkdlg__table';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    for (const label of [t.headerIndex, t.headerKind, t.headerTarget, t.headerPart]) {
      const th = document.createElement('th');
      th.textContent = label;
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    for (const link of links) {
      const row = document.createElement('tr');
      const idx = document.createElement('td');
      idx.textContent = String(link.index);
      const kind = document.createElement('td');
      kind.textContent = link.kind;
      const target = document.createElement('td');
      target.className = 'fc-extlinkdlg__cell-target';
      target.textContent = link.target || '—';
      target.title = link.target;
      const part = document.createElement('td');
      part.textContent = link.partPath;
      row.append(idx, kind, target, part);
      tbody.appendChild(row);
    }
    table.appendChild(tbody);
    tableWrap.appendChild(table);
  };

  const open = (): void => {
    renderTable();
    overlay.hidden = false;
  };

  const close = (): void => {
    overlay.hidden = true;
  };

  const refresh = (): void => {
    strings = deps.strings ?? defaultStrings;
    t = strings.externalLinksDialog;
    overlay.setAttribute('aria-label', t.title);
    header.textContent = t.title;
    note.textContent = t.note;
    empty.textContent = t.empty;
    closeBtn.textContent = t.close;
    if (!overlay.hidden) renderTable();
  };

  return {
    open,
    close,
    refresh,
    detach() {
      overlay.removeEventListener('click', onOverlayClick);
      document.removeEventListener('keydown', onKey);
      overlay.remove();
    },
  };
}
