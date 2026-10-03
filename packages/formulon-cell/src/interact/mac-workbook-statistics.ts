import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { appendDialogButton, appendDialogFrame, createDialogShell } from './dialog-shell.js';

export interface WorkbookStatistics {
  sheets: number;
  populatedCells: number;
  formulas: number;
  numbers: number;
  text: number;
  booleans: number;
  errors: number;
  tables: number;
  comments: number;
  hyperlinks: number;
  usedRows: number;
  usedColumns: number;
}

export interface MacWorkbookStatisticsDeps {
  host: HTMLElement;
  getWb: () => WorkbookHandle;
  store?: SpreadsheetStore;
  strings?: Strings;
}

export interface MacWorkbookStatisticsHandle {
  open(): void;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  detach(): void;
}

const valueIsPopulated = (cell: { value: { kind: string }; formula: string | null }): boolean =>
  cell.value.kind !== 'blank' || cell.formula !== null;

const parseStoreAddress = (key: string): { sheet: number; row: number; col: number } | null => {
  const parts = key.split(':');
  if (parts.length !== 3 || parts.some((part) => part.trim() === '')) return null;
  const values = parts.map(Number);
  if (values.some((part) => !Number.isSafeInteger(part) || part < 0)) return null;
  const [sheet, row, col] = values as [number, number, number];
  return { sheet, row, col };
};

/** Count real workbook data rather than the active selection. Physical cells
 *  are used where available so PivotTable projections do not inflate counts;
 *  every valid store-format address and native annotation anchor contributes to
 *  used-row/column bounds without expanding hyperlink rectangles.
 */
export function collectWorkbookStatistics(
  workbook: WorkbookHandle,
  store?: SpreadsheetStore,
): WorkbookStatistics {
  let populatedCells = 0;
  let formulas = 0;
  let numbers = 0;
  let text = 0;
  let booleans = 0;
  let errors = 0;
  const usedRows = new Set<string>();
  const usedColumns = new Set<string>();
  let comments = 0;
  let hyperlinks = 0;
  const addUsedAddress = (sheet: number, row: number, col: number): boolean => {
    if (
      !Number.isSafeInteger(sheet) ||
      !Number.isSafeInteger(row) ||
      !Number.isSafeInteger(col) ||
      sheet < 0 ||
      row < 0 ||
      col < 0 ||
      row > MAX_ROW ||
      col > MAX_COL ||
      sheet >= workbook.sheetCount
    ) {
      return false;
    }
    usedRows.add(`${sheet}:${row}`);
    usedColumns.add(`${sheet}:${col}`);
    return true;
  };

  for (let sheet = 0; sheet < workbook.sheetCount; sheet += 1) {
    const physical =
      typeof workbook.physicalCells === 'function'
        ? workbook.physicalCells(sheet)
        : workbook.cells(sheet);
    for (const cell of physical) {
      if (!valueIsPopulated(cell)) continue;
      populatedCells += 1;
      addUsedAddress(sheet, cell.addr.row, cell.addr.col);
      if (cell.formula !== null) formulas += 1;
      switch (cell.value.kind) {
        case 'number':
          numbers += 1;
          break;
        case 'text':
          text += 1;
          break;
        case 'bool':
          booleans += 1;
          break;
        case 'error':
          errors += 1;
          break;
      }
    }
    if (workbook.capabilities.commentsEnumerable) {
      const nativeComments = workbook.getComments(sheet);
      comments += nativeComments.length;
      for (const comment of nativeComments) addUsedAddress(sheet, comment.row, comment.col);
    }
    if (workbook.capabilities.hyperlinks) {
      const nativeHyperlinks = workbook.getHyperlinks(sheet);
      hyperlinks += nativeHyperlinks.length;
      for (const hyperlink of nativeHyperlinks) {
        // Hyperlink records may cover rectangles. Statistics use the native
        // record count and anchor cell only; do not expand the rectangle.
        addUsedAddress(sheet, hyperlink.row, hyperlink.col);
      }
    }
  }

  // Stub/older engines may not enumerate metadata. Hydrated format entries
  // provide used bounds for every formatted cell and fallback counts for
  // comment/hyperlink entries when native enumeration is unavailable.
  if (store) {
    const seenComments = new Set<string>();
    const seenHyperlinks = new Set<string>();
    for (const [key, format] of store.getState().format.formats) {
      const addr = parseStoreAddress(key);
      if (!addr) continue;
      if (!addUsedAddress(addr.sheet, addr.row, addr.col)) continue;
      const canonicalKey = `${addr.sheet}:${addr.row}:${addr.col}`;
      if (format.comment) seenComments.add(canonicalKey);
      if (format.hyperlink) seenHyperlinks.add(canonicalKey);
    }
    if (!workbook.capabilities.commentsEnumerable) comments = seenComments.size;
    if (!workbook.capabilities.hyperlinks) hyperlinks = seenHyperlinks.size;
  }

  return {
    sheets: workbook.sheetCount,
    populatedCells,
    formulas,
    numbers,
    text,
    booleans,
    errors,
    tables: workbook.getTables().length,
    comments,
    hyperlinks,
    usedRows: usedRows.size,
    usedColumns: usedColumns.size,
  };
}

const STAT_FIELDS: readonly (keyof WorkbookStatistics)[] = [
  'sheets',
  'populatedCells',
  'formulas',
  'numbers',
  'text',
  'booleans',
  'errors',
  'tables',
  'comments',
  'hyperlinks',
  'usedRows',
  'usedColumns',
];

/** A compact Review > Workbook Statistics dialog backed by the current
 *  workbook handle. Re-opening always re-counts, so it remains correct after
 *  edits and `setWorkbook()` swaps. */
export function attachMacWorkbookStatistics(
  deps: MacWorkbookStatisticsDeps,
): MacWorkbookStatisticsHandle {
  const { host } = deps;
  let strings = deps.strings ?? defaultStrings;
  const shell = createDialogShell({
    host,
    className: 'fc-macstatsdlg',
    ariaLabel: strings.macWorkbookStats.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: strings.macWorkbookStats.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-macstatsdlg__panel'],
    bodyClass: 'fc-fmtdlg__body fc-macstatsdlg__body',
  });
  const list = document.createElement('dl');
  list.className = 'fc-macstatsdlg__list';
  list.dataset.macWorkbookStatistics = 'true';
  body.appendChild(list);
  const closeBtn = appendDialogButton(footer, {
    label: strings.macWorkbookStats.close,
    variant: 'secondary',
  });

  shell.on(closeBtn, 'click', () => api.close());

  const render = (): void => {
    const stats = collectWorkbookStatistics(deps.getWb(), deps.store);
    const labels = strings.macWorkbookStats;
    list.replaceChildren();
    for (const key of STAT_FIELDS) {
      const term = document.createElement('dt');
      term.textContent = labels[key];
      term.dataset.macWorkbookStat = key;
      const value = document.createElement('dd');
      value.textContent = String(stats[key]);
      value.dataset.macWorkbookStatValue = key;
      list.append(term, value);
    }
  };

  const refreshLabels = (): void => {
    const title = strings.macWorkbookStats.title;
    header.textContent = title;
    shell.setAriaLabel(title);
    closeBtn.textContent = strings.macWorkbookStats.close;
  };

  const api: MacWorkbookStatisticsHandle = {
    open(): void {
      render();
      shell.open();
      requestAnimationFrame(() => closeBtn.focus());
    },
    close(): void {
      shell.close();
      host.focus();
    },
    refresh(): void {
      refreshLabels();
      if (shell.isOpen()) render();
    },
    setStrings(next: Strings): void {
      strings = next;
      refreshLabels();
      if (shell.isOpen()) render();
    },
    detach(): void {
      shell.dispose();
    },
  };

  return api;
}
