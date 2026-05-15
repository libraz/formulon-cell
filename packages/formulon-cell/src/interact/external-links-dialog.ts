import { listExternalLinks } from '../commands/external-links.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { createDialogShell } from './dialog-shell.js';

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
 * Spreadsheet-style "Edit Links" — read-only inventory of external-reference
 * records carried by the workbook. Records that round-trip through formulon's
 * passthrough mechanism stay listed but are not editable through this UI.
 */
export function attachExternalLinksDialog(
  deps: ExternalLinksDialogDeps,
): ExternalLinksDialogHandle {
  const { host, getWb } = deps;
  let strings = deps.strings ?? defaultStrings;
  let t = strings.externalLinksDialog;

  const shell = createDialogShell({
    host,
    className: 'fc-extlinkdlg',
    ariaLabel: t.title,
    onDismiss: () => shell.close(),
  });

  const header = document.createElement('div');
  header.className = 'fc-extlinkdlg__header';
  header.textContent = t.title;
  shell.panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-extlinkdlg__body';
  shell.panel.appendChild(body);

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
  shell.panel.appendChild(footer);

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-extlinkdlg__close';
  closeBtn.textContent = t.close;
  shell.on(closeBtn, 'click', () => shell.close());
  footer.appendChild(closeBtn);
  let selectedIndex = 0;

  const focusRow = (idx: number): void => {
    const rows = Array.from(tableWrap.querySelectorAll<HTMLTableRowElement>('tbody tr'));
    if (rows.length === 0) return;
    selectedIndex = (idx + rows.length) % rows.length;
    for (const [rowIdx, row] of rows.entries()) {
      const selected = rowIdx === selectedIndex;
      row.tabIndex = selected ? 0 : -1;
      row.setAttribute('aria-selected', selected ? 'true' : 'false');
    }
    rows[selectedIndex]?.focus({ preventScroll: true });
  };

  const renderTable = (): void => {
    tableWrap.replaceChildren();
    const links = listExternalLinks(getWb());
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
    for (const [rowIndex, link] of links.entries()) {
      const row = document.createElement('tr');
      row.tabIndex = rowIndex === selectedIndex ? 0 : -1;
      row.setAttribute('aria-selected', rowIndex === selectedIndex ? 'true' : 'false');
      row.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          focusRow(rowIndex + 1);
        } else if (e.key === 'ArrowUp') {
          e.preventDefault();
          focusRow(rowIndex - 1);
        } else if (e.key === 'Home') {
          e.preventDefault();
          focusRow(0);
        } else if (e.key === 'End') {
          e.preventDefault();
          focusRow(links.length - 1);
        }
      });
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
    selectedIndex = Math.min(selectedIndex, links.length - 1);
    focusRow(selectedIndex);
  };

  return {
    open() {
      renderTable();
      shell.open();
    },
    close() {
      shell.close();
    },
    refresh() {
      strings = deps.strings ?? defaultStrings;
      t = strings.externalLinksDialog;
      shell.setAriaLabel(t.title);
      header.textContent = t.title;
      note.textContent = t.note;
      empty.textContent = t.empty;
      closeBtn.textContent = t.close;
      if (shell.isOpen()) renderTable();
    },
    detach() {
      shell.dispose();
    },
  };
}
