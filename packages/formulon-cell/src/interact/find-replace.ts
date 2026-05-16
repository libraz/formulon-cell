import {
  applySubstitution,
  type FindMatch,
  type FindOptions,
  findAll,
  findNext,
  replaceAll,
  replaceOne,
} from '../commands/find.js';
import type { History } from '../commands/history.js';
import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

export interface FindReplaceDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings?: Strings;
  history?: History | null;
  onAfterCommit?: () => void;
}

export interface FindReplaceHandle {
  open(tab?: 'find' | 'replace'): void;
  close(): void;
  /** Swap the active dictionary; live-updates labels in place. */
  setStrings(next: Strings): void;
  detach(): void;
}

type FindTab = 'find' | 'replace';

export function attachFindReplace(deps: FindReplaceDeps): FindReplaceHandle {
  const { host, store, wb } = deps;
  let strings = deps.strings ?? defaultStrings;
  let activeTab: FindTab = 'find';
  let optionsOpen = false;

  const overlay = document.createElement('div');
  overlay.className = 'fc-find';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'false');
  overlay.hidden = true;

  const titlebar = document.createElement('div');
  titlebar.className = 'fc-find__titlebar';
  const title = document.createElement('div');
  title.className = 'fc-find__title';
  const closeBtn = makeBtn('fc-find__btn--icon');
  closeBtn.textContent = '×';
  titlebar.append(title, closeBtn);

  const tabs = document.createElement('div');
  tabs.className = 'fc-find__tabs';
  tabs.setAttribute('role', 'tablist');
  const findTab = makeBtn('fc-find__tab');
  const replaceTab = makeBtn('fc-find__tab');
  findTab.setAttribute('role', 'tab');
  replaceTab.setAttribute('role', 'tab');
  tabs.append(findTab, replaceTab);

  const form = document.createElement('div');
  form.className = 'fc-find__form';

  const findInput = document.createElement('input');
  findInput.type = 'text';
  findInput.spellcheck = false;
  findInput.autocomplete = 'off';
  const findRow = labeledControl('fc-find-find-input', findInput);

  const replaceInput = document.createElement('input');
  replaceInput.type = 'text';
  replaceInput.spellcheck = false;
  replaceInput.autocomplete = 'off';
  const replaceRow = labeledControl('fc-find-replace-input', replaceInput);
  replaceRow.row.classList.add('fc-find__replace-row');

  const optionsBtn = makeBtn('fc-find__options-btn');
  const optionsRow = document.createElement('div');
  optionsRow.className = 'fc-find__options-row';
  optionsRow.append(optionsBtn);

  const optionsPanel = document.createElement('div');
  optionsPanel.className = 'fc-find__options';

  const withinSelect = makeSelect([
    ['sheet', 'Sheet'],
    ['workbook', 'Workbook'],
  ]);
  const searchSelect = makeSelect([
    ['rows', 'By Rows'],
    ['columns', 'By Columns'],
  ]);
  const lookInSelect = makeSelect([
    ['formulas', 'Formulas'],
    ['values', 'Values'],
    ['comments', 'Comments'],
    ['notes', 'Notes'],
  ]);
  lookInSelect.value = 'values';

  const withinRow = labeledControl('fc-find-within', withinSelect, 'fc-find__field--short');
  const searchRow = labeledControl('fc-find-search', searchSelect, 'fc-find__field--short');
  const lookInRow = labeledControl('fc-find-look-in', lookInSelect, 'fc-find__field--short');

  const formatBtn = makeBtn('fc-find__format-btn');
  formatBtn.disabled = true;
  const formatRow = document.createElement('div');
  formatRow.className = 'fc-find__field fc-find__field--format';
  const formatSpacer = document.createElement('span');
  formatRow.append(formatSpacer, formatBtn);

  const caseToggle = document.createElement('input');
  caseToggle.type = 'checkbox';
  caseToggle.id = 'fc-find-case';
  const caseLabel = document.createElement('label');
  caseLabel.className = 'fc-find__check';
  const caseText = document.createElement('span');
  caseLabel.append(caseToggle, caseText);

  const wholeToggle = document.createElement('input');
  wholeToggle.type = 'checkbox';
  wholeToggle.id = 'fc-find-whole';
  const wholeLabel = document.createElement('label');
  wholeLabel.className = 'fc-find__check';
  const wholeText = document.createElement('span');
  wholeLabel.append(wholeToggle, wholeText);

  const checks = document.createElement('div');
  checks.className = 'fc-find__checks';
  checks.append(caseLabel, wholeLabel);

  optionsPanel.append(withinRow.row, searchRow.row, lookInRow.row, checks, formatRow);
  form.append(findRow.row, replaceRow.row, optionsRow, optionsPanel);

  const results = document.createElement('div');
  results.className = 'fc-find__results';
  results.hidden = true;
  results.setAttribute('aria-live', 'polite');
  const resultsTable = document.createElement('table');
  resultsTable.className = 'fc-find__results-table';
  const resultsHead = document.createElement('thead');
  const resultsHeadRow = document.createElement('tr');
  const bookHead = document.createElement('th');
  bookHead.className = 'fc-find__result-col--sheet';
  const cellHead = document.createElement('th');
  cellHead.className = 'fc-find__result-col--cell';
  const valueHead = document.createElement('th');
  resultsHeadRow.append(bookHead, cellHead, valueHead);
  resultsHead.append(resultsHeadRow);
  const resultsBody = document.createElement('tbody');
  resultsTable.append(resultsHead, resultsBody);
  const resultsSummary = document.createElement('div');
  resultsSummary.className = 'fc-find__results-summary fc-find__pill';
  resultsSummary.setAttribute('aria-live', 'polite');
  results.append(resultsTable, resultsSummary);

  const footer = document.createElement('div');
  footer.className = 'fc-find__footer';
  const findAllBtn = makeBtn();
  findAllBtn.classList.add('fc-find__btn--find-all');
  const prevBtn = makeBtn();
  prevBtn.classList.add('fc-find__btn--prev');
  const nextBtn = makeBtn();
  nextBtn.classList.add('fc-find__btn--next');
  const replaceBtn = makeBtn();
  replaceBtn.classList.add('fc-find__btn--replace');
  const replaceAllBtn = makeBtn();
  replaceAllBtn.classList.add('fc-find__btn--replace-all');
  const closeFooterBtn = makeBtn();
  closeFooterBtn.classList.add('fc-find__btn--close');
  footer.append(findAllBtn, prevBtn, nextBtn, replaceBtn, replaceAllBtn, closeFooterBtn);

  overlay.append(titlebar, tabs, form, results, footer);
  host.appendChild(overlay);

  const relabel = (): void => {
    const t = strings.findReplace;
    overlay.setAttribute('aria-label', t.title);
    title.textContent = t.title;
    findTab.textContent = t.findTab;
    replaceTab.textContent = t.replaceTab;
    findRow.label.textContent = t.findWhat;
    findInput.setAttribute('aria-label', t.findWhat);
    replaceRow.label.textContent = t.replaceWith;
    replaceInput.setAttribute('aria-label', t.replaceWith);
    optionsBtn.textContent = optionsOpen ? t.optionsLess : t.optionsMore;
    withinRow.label.textContent = t.within;
    searchRow.label.textContent = t.search;
    lookInRow.label.textContent = t.lookIn;
    setOptions(withinSelect, [t.sheet, t.workbook]);
    setOptions(searchSelect, [t.byRows, t.byColumns]);
    setOptions(lookInSelect, [t.formulas, t.values, t.comments, t.notes]);
    caseText.textContent = t.matchCase;
    wholeText.textContent = t.matchEntire;
    formatBtn.textContent = t.format;
    findAllBtn.textContent = t.findAll;
    prevBtn.textContent = t.prev;
    nextBtn.textContent = t.next;
    replaceBtn.textContent = t.replaceOne;
    replaceAllBtn.textContent = t.replaceAll;
    closeBtn.setAttribute('aria-label', t.close);
    closeBtn.title = t.close;
    closeFooterBtn.textContent = t.close;
    bookHead.textContent = t.bookHeader;
    cellHead.textContent = t.cellHeader;
    valueHead.textContent = t.valueHeader;
    syncTab();
  };

  let lastQuery = '';
  let currentMatch: FindMatch | null = null;

  const opts = (): FindOptions => ({
    query: findInput.value,
    caseSensitive: caseToggle.checked,
    matchWhole: wholeToggle.checked,
    within: withinSelect.value === 'workbook' ? 'workbook' : 'sheet',
    searchBy: searchSelect.value === 'columns' ? 'columns' : 'rows',
    lookIn:
      lookInSelect.value === 'formulas' ||
      lookInSelect.value === 'comments' ||
      lookInSelect.value === 'notes'
        ? lookInSelect.value
        : 'values',
  });

  const clearResults = (): void => {
    resultsBody.replaceChildren();
    results.hidden = true;
    resultsSummary.textContent = '';
  };

  const renderResults = (): void => {
    const t = strings.findReplace;
    const all = findAll(store.getState(), opts());
    resultsBody.replaceChildren();
    for (const match of all) {
      const tr = document.createElement('tr');
      tr.tabIndex = 0;
      const sheetName = wb.sheetName(match.addr.sheet);
      const value = valueForMatch(match.addr, opts().lookIn ?? 'values');
      tr.append(td(sheetName), td(addrLabel(match.addr)), td(value));
      tr.addEventListener('click', () => {
        currentMatch = match;
        mutators.setActive(store, match.addr);
        updateSummary(all.length);
      });
      tr.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          tr.click();
        }
      });
      resultsBody.append(tr);
    }
    results.hidden = false;
    updateSummary(all.length);
    resultsSummary.textContent = t.cellsFound.replace('{count}', String(all.length));
  };

  const updateSummary = (count?: number, text?: string): void => {
    if (text !== undefined) {
      resultsSummary.textContent = text;
      return;
    }
    const allCount = count ?? findAll(store.getState(), opts()).length;
    if (!opts().query) {
      resultsSummary.textContent = '0 / 0';
      return;
    }
    const idx =
      currentMatch === null
        ? 0
        : findAll(store.getState(), opts()).findIndex(
            (m) =>
              m.addr.sheet === currentMatch?.addr.sheet &&
              m.addr.row === currentMatch.addr.row &&
              m.addr.col === currentMatch.addr.col,
          ) + 1;
    resultsSummary.textContent = `${idx > 0 ? idx : 0} / ${allCount}`;
  };

  const step = (direction: 'next' | 'prev'): void => {
    const o = opts();
    if (!o.query) {
      currentMatch = null;
      updateSummary();
      return;
    }
    const state = store.getState();
    const from = currentMatch ? currentMatch.addr : state.selection.active;
    const m = findNext(state, o, from, direction);
    currentMatch = m;
    if (m) mutators.setActive(store, m.addr);
    updateSummary();
  };

  const doReplace = (): void => {
    const o = opts();
    if (!o.query) {
      currentMatch = null;
      updateSummary();
      return;
    }
    if (!currentMatch) {
      const state = store.getState();
      const active = state.selection.active;
      const matches = findAll(state, o);
      currentMatch =
        matches.find(
          (m) =>
            m.addr.sheet === active.sheet && m.addr.row === active.row && m.addr.col === active.col,
        ) ??
        findNext(state, o, active, 'next') ??
        null;
      if (currentMatch) mutators.setActive(store, currentMatch.addr);
      else {
        updateSummary();
        return;
      }
    }
    if (wb.cellFormula(currentMatch.addr) === null) {
      const cur = formatCell(wb.getValue(currentMatch.addr));
      const next = applySubstitution(cur, o, replaceInput.value);
      if (next !== cur) {
        if (replaceOne(wb, currentMatch, next, store)) deps.onAfterCommit?.();
      }
    }
    clearResults();
    step('next');
  };

  const doReplaceAll = (): void => {
    const o = opts();
    if (!o.query) {
      updateSummary();
      return;
    }
    const history = deps.history ?? null;
    if (history) history.begin();
    let n = 0;
    try {
      n = replaceAll(store.getState(), wb, o, replaceInput.value, store);
    } finally {
      if (history) history.end();
    }
    if (n > 0) deps.onAfterCommit?.();
    currentMatch = null;
    clearResults();
    updateSummary(undefined, strings.findReplace.replacedCount.replace('{count}', String(n)));
  };

  const syncTab = (): void => {
    const t = strings.findReplace;
    findTab.setAttribute('aria-selected', activeTab === 'find' ? 'true' : 'false');
    replaceTab.setAttribute('aria-selected', activeTab === 'replace' ? 'true' : 'false');
    replaceRow.row.hidden = activeTab === 'find';
    replaceBtn.hidden = activeTab === 'find';
    replaceAllBtn.hidden = activeTab === 'find';
    findAllBtn.hidden = activeTab === 'replace';
    lookInSelect.disabled = activeTab === 'replace';
    if (activeTab === 'replace') lookInSelect.value = 'formulas';
    findInput.placeholder = activeTab === 'find' ? t.findLabel : t.findWhat;
  };

  const showTab = (tab: FindTab): void => {
    activeTab = tab;
    currentMatch = null;
    clearResults();
    syncTab();
    updateSummary();
    findInput.focus();
  };

  const syncOptions = (): void => {
    optionsPanel.hidden = !optionsOpen;
    optionsBtn.textContent = optionsOpen
      ? strings.findReplace.optionsLess
      : strings.findReplace.optionsMore;
  };

  const resetSearch = (): void => {
    currentMatch = null;
    clearResults();
    updateSummary();
  };

  const onFindKey = (e: KeyboardEvent): void => {
    if (e.key === 'Enter') {
      e.preventDefault();
      step(e.shiftKey ? 'prev' : 'next');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  const onReplaceKey = (e: KeyboardEvent): void => {
    if (e.key === 'Enter') {
      e.preventDefault();
      doReplace();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  const onPrevClick = (): void => step('prev');
  const onNextClick = (): void => step('next');
  const onCloseClick = (): void => api.close();
  const onOptionsClick = (): void => {
    optionsOpen = !optionsOpen;
    syncOptions();
  };

  findTab.addEventListener('click', () => showTab('find'));
  replaceTab.addEventListener('click', () => showTab('replace'));
  optionsBtn.addEventListener('click', onOptionsClick);
  findAllBtn.addEventListener('click', renderResults);
  prevBtn.addEventListener('click', onPrevClick);
  nextBtn.addEventListener('click', onNextClick);
  replaceBtn.addEventListener('click', doReplace);
  replaceAllBtn.addEventListener('click', doReplaceAll);
  closeBtn.addEventListener('click', onCloseClick);
  closeFooterBtn.addEventListener('click', onCloseClick);
  findInput.addEventListener('keydown', onFindKey);
  findInput.addEventListener('input', resetSearch);
  replaceInput.addEventListener('keydown', onReplaceKey);
  for (const control of [caseToggle, wholeToggle, withinSelect, searchSelect, lookInSelect]) {
    control.addEventListener('change', resetSearch);
  }
  overlay.addEventListener('keydown', onOverlayKey);

  const valueForMatch = (addr: Addr, lookIn: FindOptions['lookIn']): string => {
    const key = addrKey(addr);
    const cell = store.getState().data.cells.get(key);
    if (lookIn === 'formulas') return cell ? (cell.formula ?? formatCell(cell.value)) : '';
    if (lookIn === 'comments' || lookIn === 'notes') {
      return store.getState().format.formats.get(key)?.comment ?? '';
    }
    return cell ? formatCell(cell.value) : '';
  };

  const api: FindReplaceHandle = {
    open(tab: FindTab = 'find'): void {
      overlay.hidden = false;
      activeTab = tab;
      findInput.value = lastQuery;
      currentMatch = null;
      clearResults();
      syncOptions();
      syncTab();
      updateSummary();
      requestAnimationFrame(() => {
        findInput.focus();
        findInput.select();
      });
    },
    close(): void {
      lastQuery = findInput.value;
      overlay.hidden = true;
      currentMatch = null;
      clearResults();
      host.focus();
    },
    setStrings(next: Strings): void {
      strings = next;
      relabel();
    },
    detach(): void {
      optionsBtn.removeEventListener('click', onOptionsClick);
      findAllBtn.removeEventListener('click', renderResults);
      prevBtn.removeEventListener('click', onPrevClick);
      nextBtn.removeEventListener('click', onNextClick);
      replaceBtn.removeEventListener('click', doReplace);
      replaceAllBtn.removeEventListener('click', doReplaceAll);
      closeBtn.removeEventListener('click', onCloseClick);
      closeFooterBtn.removeEventListener('click', onCloseClick);
      findInput.removeEventListener('keydown', onFindKey);
      findInput.removeEventListener('input', resetSearch);
      replaceInput.removeEventListener('keydown', onReplaceKey);
      for (const control of [caseToggle, wholeToggle, withinSelect, searchSelect, lookInSelect]) {
        control.removeEventListener('change', resetSearch);
      }
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };

  relabel();
  syncOptions();

  return api;
}

function makeBtn(extraClass?: string): HTMLButtonElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = extraClass ? `fc-find__btn ${extraClass}` : 'fc-find__btn';
  return b;
}

function makeSelect(options: [string, string][]): HTMLSelectElement {
  const select = document.createElement('select');
  for (const [value, text] of options) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = text;
    select.append(option);
  }
  return select;
}

function setOptions(select: HTMLSelectElement, labels: string[]): void {
  labels.forEach((label, index) => {
    const option = select.options.item(index);
    if (option) option.textContent = label;
  });
}

function labeledControl<T extends HTMLElement>(
  id: string,
  control: T,
  extraClass?: string,
): { row: HTMLDivElement; label: HTMLLabelElement; control: T } {
  const row = document.createElement('div');
  row.className = extraClass ? `fc-find__field ${extraClass}` : 'fc-find__field';
  const label = document.createElement('label');
  label.htmlFor = id;
  control.id = id;
  row.append(label, control);
  return { row, label, control };
}

function td(text: string): HTMLTableCellElement {
  const cell = document.createElement('td');
  cell.textContent = text;
  return cell;
}

function addrLabel(addr: Addr): string {
  return `${colLabel(addr.col)}${addr.row + 1}`;
}

function colLabel(col: number): string {
  let n = col + 1;
  let label = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    label = String.fromCharCode(65 + r) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}
