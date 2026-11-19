import {
  applySubstitution,
  displayTextForFindMatch,
  type FindMatch,
  type FindOptions,
  findAll,
  findDisplayResolver,
  replaceAll,
  replaceOne,
} from '../commands/find.js';
import { type History, transactIn } from '../commands/history.js';
import { deferForAllowedRangeUnlock } from '../commands/protection.js';
import { addrKey, formatA1Cell } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { sameAddr } from '../store/direct-format.js';
import { rangeOfAddr } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { createDialogSelect, type DialogSelectOption } from '../toolbar/dialogs/form-controls.js';
import { showMessage } from '../toolbar/dialogs/prompt.js';
import { projectDisabledReason, projectDisabledState } from '../toolbar/menu-a11y.js';
import { createDialogButton, markSurface, openerOf, returnFocus } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';
import { createMacFindPanel, type MacFindPanel } from './mac-find-panel.js';
import { isNavigationAddrAllowed, navigationPolicyFor } from './navigation-policy.js';
import { openPopupMenu } from './popup-menu.js';
import { bindRovingRows, setRovingStop } from './roving-rows.js';

export interface FindReplaceDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings?: Strings;
  /** Current UI locale. The workbook profile remains the fallback for
   * low-level hosts that do not provide this getter. */
  getLocale: () => string;
  /** Activate and hydrate a workbook sheet before selecting a match on it.
   * Return false when the sheet is hidden or blocked by a host policy. */
  onActivateSheet?: (sheet: number) => boolean;
  history?: History | null;
  onAfterCommit?: () => void;
}

export interface FindReplaceHandle {
  /** Opens the dialog on `tab`. With no tab, opens the host's Find: the floating find panel on
   *  the Mac, the dialog's Find tab elsewhere. */
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
  const isMac = (): boolean => isMacHost(host);
  /** The dialog's strings, with the Mac wording laid over them on the Mac host. */
  const copy = (): Strings['findReplace'] =>
    isMac() ? { ...strings.findReplace, ...strings.mac.findReplace } : strings.findReplace;
  const widthText = (): string => strings.mac.findReplace.matchWidth;
  let activeTab: FindTab = 'find';
  let optionsOpen = false;

  const overlay = document.createElement('div');
  let opener: HTMLElement | null = null;
  overlay.className = 'fc-find fc-fmtdlg__panel';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'false');
  markSurface(overlay, 'floating');
  overlay.hidden = true;

  const titlebar = document.createElement('div');
  titlebar.className = 'fc-fmtdlg__header fc-find__titlebar';
  const title = document.createElement('div');
  title.className = 'fc-find__title';
  const closeBtn = createDialogButton({
    label: '',
    baseClass: 'fc-fmtdlg__close fc-find__btn--icon',
  });
  titlebar.append(title, closeBtn);

  const tabs = document.createElement('div');
  tabs.className = 'fc-fmtdlg__tabs fc-find__tabs';
  tabs.setAttribute('role', 'tablist');
  const findTab = createDialogButton({ label: '', baseClass: 'fc-fmtdlg__tab fc-find__tab' });
  const replaceTab = createDialogButton({ label: '', baseClass: 'fc-fmtdlg__tab fc-find__tab' });
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
  // The Mac find field drops down this session's searches, most recent first.
  const historyBtn = makeBtn('fc-find__history');
  historyBtn.setAttribute('aria-haspopup', 'menu');
  historyBtn.setAttribute('aria-expanded', 'false');
  const findCombo = document.createElement('div');
  findCombo.className = 'fc-find__combo';
  findRow.row.replaceChild(findCombo, findInput);
  findCombo.append(findInput);
  if (isMac()) findCombo.append(historyBtn);
  const searches: string[] = [];

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
  lookInSelect.value = isMac() ? 'formulas' : 'values';

  const withinRow = labeledControl('fc-find-within', withinSelect, 'fc-find__field--short');
  const searchRow = labeledControl('fc-find-search', searchSelect, 'fc-find__field--short');
  const lookInRow = labeledControl('fc-find-look-in', lookInSelect, 'fc-find__field--short');

  const formatBtn = makeBtn('fc-find__format-btn');
  projectDisabledState(formatBtn, true, copy().formatUnavailable, {
    datasetKey: 'disabledReason',
    titlePrefix: copy().format,
  });
  const formatRow = document.createElement('div');
  formatRow.className = 'fc-find__field fc-find__field--format';
  const formatSpacer = document.createElement('span');
  formatRow.append(formatSpacer, formatBtn);
  formatRow.hidden = isMac();

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

  const widthToggle = document.createElement('input');
  widthToggle.type = 'checkbox';
  widthToggle.id = 'fc-find-width';
  const widthLabel = document.createElement('label');
  widthLabel.className = 'fc-find__check';
  const widthText_ = document.createElement('span');
  widthLabel.append(widthToggle, widthText_);

  const checks = document.createElement('div');
  checks.className = 'fc-find__checks';
  checks.append(caseLabel, wholeLabel);
  // The half-width / full-width option exists on the Mac host only.
  if (isMac()) checks.append(widthLabel);

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
  const resultRows = (): HTMLElement[] =>
    Array.from(resultsBody.querySelectorAll<HTMLElement>(':scope > tr'));
  // The results take one tab stop; the arrow keys move between them.
  bindRovingRows(resultsBody, resultRows);
  const resultsSummary = document.createElement('div');
  resultsSummary.className = 'fc-find__results-summary fc-find__pill';
  resultsSummary.setAttribute('aria-live', 'polite');
  results.append(resultsTable, resultsSummary);

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer fc-find__footer';
  const findAllBtn = makeBtn('fc-find__btn--find-all');
  const prevBtn = makeBtn('fc-find__btn--prev');
  const nextBtn = makeBtn('fc-find__btn--next');
  const replaceBtn = makeBtn('fc-find__btn--replace');
  const replaceAllBtn = makeBtn('fc-find__btn--replace-all');
  const closeFooterBtn = makeBtn('fc-find__btn--close');
  footer.append(findAllBtn, prevBtn, nextBtn, replaceBtn, replaceAllBtn, closeFooterBtn);

  // The form and results scroll together when the host is shorter than the panel.
  const body = document.createElement('div');
  body.className = 'fc-find__body';
  body.append(form, results);

  // The Mac app keeps a cell count under the buttons.
  const status = document.createElement('div');
  status.className = 'fc-find__status';
  status.setAttribute('aria-live', 'polite');
  status.hidden = !isMac();
  const setStatus = (count: number): void => {
    status.textContent = copy().cellsFound.replace('{count}', String(count));
  };

  footer.append(status);
  overlay.append(titlebar, tabs, body, footer);
  host.appendChild(overlay);

  const relabel = (): void => {
    const t = copy();
    overlay.setAttribute('aria-label', t.title);
    title.textContent = t.title;
    findTab.textContent = t.findTab;
    replaceTab.textContent = t.replaceTab;
    findRow.label.textContent = t.findWhat;
    findInput.setAttribute('aria-label', t.findWhat);
    historyBtn.setAttribute('aria-label', t.findWhat);
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
    widthText_.textContent = widthText();
    formatBtn.textContent = t.format;
    projectDisabledState(formatBtn, true, t.formatUnavailable, {
      datasetKey: 'disabledReason',
      titlePrefix: t.format,
    });
    findAllBtn.textContent = t.findAll;
    prevBtn.textContent = t.prev;
    nextBtn.textContent = t.next;
    // Both step buttons are primary on the Mac.
    prevBtn.classList.toggle('fc-fmtdlg__btn--primary', isMac());
    nextBtn.classList.toggle('fc-fmtdlg__btn--primary', isMac());
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

  const remember = (query: string): void => {
    if (!query) return;
    const at = searches.indexOf(query);
    if (at >= 0) searches.splice(at, 1);
    searches.unshift(query);
    syncHistory();
  };
  const syncHistory = (): void => {
    projectDisabledState(historyBtn, searches.length === 0, null);
  };
  let closeHistory: (() => void) | null = null;
  const onHistoryClick = (): void => {
    const wasOpen = historyBtn.getAttribute('aria-expanded') === 'true';
    closeHistory?.();
    closeHistory = null;
    if (wasOpen || searches.length === 0) return;
    closeHistory = openPopupMenu({
      anchor: historyBtn,
      label: copy().findWhat,
      items: searches.map((query) => ({
        label: query,
        onPick: () => {
          findInput.value = query;
          resetSearch();
          findInput.focus();
          findInput.select();
        },
      })),
    });
  };

  const opts = (): FindOptions => ({
    query: findInput.value,
    caseSensitive: caseToggle.checked,
    matchWhole: wholeToggle.checked,
    ...(isMac() ? { matchWidth: widthToggle.checked } : {}),
    within: withinSelect.value === 'workbook' ? 'workbook' : 'sheet',
    searchBy: searchSelect.value === 'columns' ? 'columns' : 'rows',
    lookIn:
      lookInSelect.value === 'formulas' ||
      lookInSelect.value === 'comments' ||
      lookInSelect.value === 'notes'
        ? lookInSelect.value
        : 'values',
    workbook: wb,
    locale: deps.getLocale(),
  });

  const navigationRestricted = (): boolean => {
    const options = navigationPolicyFor(store)?.options;
    return !!(options?.range || options?.selectable);
  };

  const matchesForNavigation = (state: ReturnType<SpreadsheetStore['getState']>, o: FindOptions) =>
    findAll(state, o).filter(
      (match) =>
        !state.layout.hiddenSheets.has(match.addr.sheet) &&
        !state.layout.veryHiddenSheets.has(match.addr.sheet) &&
        isNavigationAddrAllowed(store, match.addr),
    );

  const nextMatchForNavigation = (
    matches: FindMatch[],
    o: FindOptions,
    from: Addr | null,
    direction: 'next' | 'prev',
  ): FindMatch | null => {
    if (matches.length === 0) return null;
    if (!from) return direction === 'next' ? (matches[0] ?? null) : (matches.at(-1) ?? null);
    const byRows = o.searchBy !== 'columns';
    const key = (addr: Addr): number =>
      addr.sheet * 1_000_000_000_000 +
      (byRows ? addr.row * 1_000_000 + addr.col : addr.col * 1_000_000 + addr.row);
    const fromKey = key(from);
    if (direction === 'next') {
      return matches.find((match) => key(match.addr) > fromKey) ?? matches[0] ?? null;
    }
    for (let i = matches.length - 1; i >= 0; i -= 1) {
      const match = matches[i];
      if (match && key(match.addr) < fromKey) return match;
    }
    return matches.at(-1) ?? null;
  };

  const activateMatch = (match: FindMatch): boolean => {
    const current = store.getState().data.sheetIndex;
    if (current !== match.addr.sheet) {
      if (
        store.getState().layout.hiddenSheets.has(match.addr.sheet) ||
        store.getState().layout.veryHiddenSheets.has(match.addr.sheet)
      )
        return false;
      if (deps.onActivateSheet && !deps.onActivateSheet(match.addr.sheet)) return false;
      if (!deps.onActivateSheet) {
        mutators.setSheetIndex(store, match.addr.sheet);
        if (store.getState().data.sheetIndex === match.addr.sheet) {
          mutators.replaceCells(store, wb.cells(match.addr.sheet));
        }
      }
      if (store.getState().data.sheetIndex !== match.addr.sheet) return false;
    }
    mutators.setActive(store, match.addr);
    const landed = sameAddr(store.getState().selection.active, match.addr);
    if (landed) mutators.setFindHighlight(store, match.addr);
    return landed;
  };

  const clearResults = (): void => {
    resultsBody.replaceChildren();
    results.hidden = true;
    resultsSummary.textContent = '';
  };

  const renderResults = (): void => {
    const t = copy();
    remember(findInput.value);
    const all = matchesForNavigation(store.getState(), { ...opts(), listFromActive: true });
    resultsBody.replaceChildren();
    const displayOf = findDisplayResolver(store.getState(), opts());
    for (const match of all) {
      const tr = document.createElement('tr');
      const sheetName = wb.sheetName(match.addr.sheet);
      const value = displayOf(match.addr);
      tr.append(td(sheetName), td(addrLabel(match.addr)), td(value));
      tr.addEventListener('click', () => {
        if (activateMatch(match)) {
          currentMatch = match;
          updateSummary(all);
        }
      });
      tr.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          tr.click();
        }
      });
      resultsBody.append(tr);
    }
    setRovingStop(resultRows());
    results.hidden = false;
    setStatus(all.length);
    resultsSummary.textContent = t.cellsFound.replace('{count}', String(all.length));
  };

  /** `matches` is the caller's already-computed navigation list, when it has one. */
  const updateSummary = (matches?: FindMatch[], text?: string): void => {
    if (text !== undefined) {
      resultsSummary.textContent = text;
      return;
    }
    const o = opts();
    if (!o.query) {
      setStatus(0);
      resultsSummary.textContent = '0 / 0';
      return;
    }
    const all = matches ?? matchesForNavigation(store.getState(), o);
    setStatus(all.length);
    const cur = currentMatch;
    const idx = cur === null ? 0 : all.findIndex((m) => sameAddr(m.addr, cur.addr)) + 1;
    resultsSummary.textContent = `${Math.max(idx, 0)} / ${all.length}`;
  };

  /** Moves to the next or previous match; true when it landed on one. */
  const step = (direction: 'next' | 'prev'): boolean => {
    const o = opts();
    if (!o.query) {
      currentMatch = null;
      updateSummary();
      return false;
    }
    remember(o.query);
    const matches = matchesForNavigation(store.getState(), o);
    let from = currentMatch ? currentMatch.addr : store.getState().selection.active;
    const attempted = new Set<string>();
    let m = nextMatchForNavigation(matches, o, from, direction);
    while (m) {
      const key = addrKey(m.addr);
      if (attempted.has(key)) break;
      attempted.add(key);
      if (activateMatch(m)) break;
      from = m.addr;
      m = nextMatchForNavigation(matches, o, from, direction);
    }
    currentMatch = m;
    updateSummary(matches);
    if (!m && isMac()) {
      const t = strings.mac.alerts.findNotFound;
      void showMessage({
        host,
        title: t.headline,
        headline: t.headline,
        message: t.message,
        okLabel: t.confirm,
        alert: true,
      });
    }
    return m !== null;
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
      const matches = matchesForNavigation(state, o);
      currentMatch =
        matches.find((m) => sameAddr(m.addr, active)) ??
        nextMatchForNavigation(matches, o, active, 'next') ??
        null;
      if (!currentMatch) {
        updateSummary();
        return;
      }
    }
    if (!activateMatch(currentMatch)) {
      currentMatch = null;
      clearResults();
      updateSummary();
      return;
    }
    if (o.lookIn !== 'comments' && o.lookIn !== 'notes') {
      if (o.lookIn === 'formulas' || wb.cellFormula(currentMatch.addr) === null) {
        const cur = displayTextForFindMatch(store.getState(), currentMatch.addr, o);
        const next = applySubstitution(cur, o, replaceInput.value);
        if (next !== cur) {
          const match = currentMatch;
          const apply = (): void => {
            if (replaceOne(wb, match, next, store, o.locale, o.lookIn)) deps.onAfterCommit?.();
          };
          if (!deferForAllowedRangeUnlock(store, [rangeOfAddr(match.addr)], apply)) apply();
        }
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
    remember(o.query);
    const history = deps.history ?? null;
    const restricted = navigationRestricted();
    const targets = findAll(store.getState(), o).map((m) => rangeOfAddr(m.addr));
    if (deferForAllowedRangeUnlock(store, targets, doReplaceAll)) return;
    let n = 0;
    transactIn(history, () => {
      if (!restricted) {
        n = replaceAll(store.getState(), wb, o, replaceInput.value, store);
      } else if (o.lookIn === 'comments' || o.lookIn === 'notes') {
        n = 0;
      } else {
        const matches = matchesForNavigation(store.getState(), o);
        wb.withBatchedRecalc(() => {
          for (const match of matches) {
            if (o.lookIn !== 'formulas' && wb.cellFormula(match.addr) !== null) continue;
            const current = displayTextForFindMatch(store.getState(), match.addr, o);
            const next = applySubstitution(current, o, replaceInput.value);
            if (next !== current && replaceOne(wb, match, next, store, o.locale, o.lookIn)) n += 1;
          }
        });
      }
    });
    if (n > 0) deps.onAfterCommit?.();
    currentMatch = null;
    clearResults();
    updateSummary(undefined, copy().replacedCount.replace('{count}', String(n)));
  };

  const syncTab = (): void => {
    const t = copy();
    findTab.setAttribute('aria-selected', activeTab === 'find' ? 'true' : 'false');
    replaceTab.setAttribute('aria-selected', activeTab === 'replace' ? 'true' : 'false');
    overlay.dataset.fcFindTab = activeTab;
    replaceRow.row.hidden = activeTab === 'find';
    replaceBtn.hidden = activeTab === 'find';
    replaceAllBtn.hidden = activeTab === 'find';
    // The Mac replace tab keeps Find All, laid out after the replace buttons.
    findAllBtn.hidden = activeTab === 'replace' && !isMac();
    if (isMac()) {
      footer.append(
        replaceBtn,
        replaceAllBtn,
        findAllBtn,
        prevBtn,
        nextBtn,
        closeFooterBtn,
        status,
      );
    }
    // Find and Replace share the same projection.  Values remain the default
    // for a new replacement session, while Formula search stays available for
    // callers that explicitly choose it; both modes preserve their source
    // semantics through the command layer.
    if (
      activeTab === 'replace' &&
      lookInSelect.value !== 'formulas' &&
      lookInSelect.value !== 'values'
    ) {
      lookInSelect.value = 'values';
    }
    projectDisabledState(lookInSelect, false, null, {
      datasetKey: 'disabledReason',
      titlePrefix: t.lookIn,
    });
    projectDisabledReason(lookInRow.label, null, {
      ariaDescription: false,
      titlePrefix: t.lookIn,
    });
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
    overlay.classList.toggle('fc-find--options', optionsOpen);
    optionsBtn.textContent = optionsOpen ? copy().optionsLess : copy().optionsMore;
  };

  const resetSearch = (): void => {
    currentMatch = null;
    clearResults();
    updateSummary();
  };

  const onFindKey = (e: KeyboardEvent): void => {
    if (isSubmitEnter(e)) {
      e.preventDefault();
      step(e.shiftKey ? 'prev' : 'next');
    }
  };

  const onReplaceKey = (e: KeyboardEvent): void => {
    if (isSubmitEnter(e)) {
      e.preventDefault();
      doReplace();
    }
  };

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  const onPrevClick = (): void => {
    step('prev');
  };
  const onNextClick = (): void => {
    step('next');
  };
  const onCloseClick = (): void => api.close();
  const onOptionsClick = (): void => {
    optionsOpen = !optionsOpen;
    syncOptions();
  };

  findTab.addEventListener('click', () => showTab('find'));
  replaceTab.addEventListener('click', () => showTab('replace'));
  optionsBtn.addEventListener('click', onOptionsClick);
  findAllBtn.addEventListener('click', renderResults);
  historyBtn.addEventListener('click', onHistoryClick);
  prevBtn.addEventListener('click', onPrevClick);
  nextBtn.addEventListener('click', onNextClick);
  replaceBtn.addEventListener('click', doReplace);
  replaceAllBtn.addEventListener('click', doReplaceAll);
  closeBtn.addEventListener('click', onCloseClick);
  closeFooterBtn.addEventListener('click', onCloseClick);
  findInput.addEventListener('keydown', onFindKey);
  findInput.addEventListener('input', resetSearch);
  replaceInput.addEventListener('keydown', onReplaceKey);
  for (const control of [
    caseToggle,
    wholeToggle,
    widthToggle,
    withinSelect,
    searchSelect,
    lookInSelect,
  ]) {
    control.addEventListener('change', resetSearch);
  }
  overlay.addEventListener('keydown', onOverlayKey);

  let panel: MacFindPanel | null = null;
  const findPanel = (): MacFindPanel => {
    panel ??= createMacFindPanel({
      host,
      strings: () => strings.mac.findPanel,
      closeLabel: () => copy().close,
      query: () => findInput.value,
      setQuery: (text) => {
        lastQuery = text;
        if (findInput.value === text) return;
        findInput.value = text;
        resetSearch();
      },
      scope: () => (withinSelect.value === 'workbook' ? 'workbook' : 'sheet'),
      setScope: (scope) => {
        withinSelect.value = scope;
        resetSearch();
      },
      step,
      openDialog: (tab) => api.open(tab),
      onClose: () => mutators.setFindHighlight(store, null),
    });
    return panel;
  };

  const api: FindReplaceHandle = {
    open(requested?: FindTab): void {
      if (requested === undefined && isMac()) {
        findPanel().open();
        return;
      }
      const tab = requested ?? 'find';
      if (!overlay.contains(document.activeElement)) opener = openerOf(document);
      overlay.hidden = false;
      activeTab = tab;
      findInput.value = lastQuery;
      currentMatch = null;
      clearResults();
      syncOptions();
      syncTab();
      updateSummary();
      findInput.focus();
      findInput.select();
    },
    close(): void {
      lastQuery = findInput.value;
      overlay.hidden = true;
      mutators.setFindHighlight(store, null);
      currentMatch = null;
      clearResults();
      const target = opener;
      opener = null;
      returnFocus(target, overlay, host);
    },
    setStrings(next: Strings): void {
      strings = next;
      relabel();
      panel?.relabel();
    },
    detach(): void {
      optionsBtn.removeEventListener('click', onOptionsClick);
      findAllBtn.removeEventListener('click', renderResults);
      historyBtn.removeEventListener('click', onHistoryClick);
      closeHistory?.();
      panel?.detach();
      prevBtn.removeEventListener('click', onPrevClick);
      nextBtn.removeEventListener('click', onNextClick);
      replaceBtn.removeEventListener('click', doReplace);
      replaceAllBtn.removeEventListener('click', doReplaceAll);
      closeBtn.removeEventListener('click', onCloseClick);
      closeFooterBtn.removeEventListener('click', onCloseClick);
      findInput.removeEventListener('keydown', onFindKey);
      findInput.removeEventListener('input', resetSearch);
      replaceInput.removeEventListener('keydown', onReplaceKey);
      for (const control of [
        caseToggle,
        wholeToggle,
        widthToggle,
        withinSelect,
        searchSelect,
        lookInSelect,
      ]) {
        control.removeEventListener('change', resetSearch);
      }
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };

  relabel();
  syncOptions();
  syncHistory();

  return api;
}

function makeBtn(extraClass: string): HTMLButtonElement {
  return createDialogButton({ label: '', baseClass: `fc-fmtdlg__btn fc-find__btn ${extraClass}` });
}

function makeSelect(options: [string, string][]): HTMLSelectElement {
  return createDialogSelect(
    options.map(([value, label]): DialogSelectOption => ({ value, label })),
    options[0]?.[0] ?? '',
    { className: '' },
  );
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
  return formatA1Cell(addr.row, addr.col);
}
