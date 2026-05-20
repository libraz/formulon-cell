import { FUNCTION_SIGNATURES } from '../commands/refs.js';
import { defaultStrings, en as enStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import {
  appendDialogSelectOptions,
  createDialogSelect,
} from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  createDialogShell,
} from './dialog-shell.js';

/** Heuristic locale detector — we don't get a `Locale` flag through deps so
 *  we sniff the active dictionary's title against the canonical English one.
 *  Cheap, and correct for the two built-in locales. Custom locales fall back
 *  to English descriptions, which is the standard desktop spreadsheets behaviour for
 *  unsupported tongues. */
const detectLocale = (s: Strings): 'en' | 'ja' =>
  s.fxDialog.title === enStrings.fxDialog.title ? 'en' : 'ja';

type FunctionCategory =
  | 'all'
  | 'recent'
  | 'logical'
  | 'lookup'
  | 'text'
  | 'datetime'
  | 'math'
  | 'financial'
  | 'dynamicArray';

const FUNCTION_CATEGORY_NAMES: Record<Exclude<FunctionCategory, 'all' | 'recent'>, readonly string[]> = {
  logical: ['IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'TRUE', 'FALSE'],
  lookup: [
    'VLOOKUP',
    'HLOOKUP',
    'XLOOKUP',
    'INDEX',
    'MATCH',
    'XMATCH',
    'OFFSET',
    'INDIRECT',
    'CHOOSE',
    'ROW',
    'COLUMN',
    'ROWS',
    'COLUMNS',
  ],
  text: [
    'CONCATENATE',
    'CONCAT',
    'TEXTJOIN',
    'TEXTSPLIT',
    'TEXTBEFORE',
    'TEXTAFTER',
    'LEFT',
    'RIGHT',
    'MID',
    'LEN',
    'UPPER',
    'LOWER',
    'PROPER',
    'TRIM',
    'SUBSTITUTE',
    'REPLACE',
    'FIND',
    'SEARCH',
    'TEXT',
    'VALUE',
    'NUMBERVALUE',
  ],
  datetime: [
    'TODAY',
    'NOW',
    'DATE',
    'YEAR',
    'MONTH',
    'DAY',
    'HOUR',
    'MINUTE',
    'SECOND',
    'WEEKDAY',
    'EOMONTH',
    'DATEDIF',
    'NETWORKDAYS',
    'WORKDAY',
  ],
  math: [
    'SUM',
    'AVERAGE',
    'COUNT',
    'COUNTA',
    'COUNTIF',
    'COUNTIFS',
    'SUMIF',
    'SUMIFS',
    'AVERAGEIF',
    'AVERAGEIFS',
    'MIN',
    'MAX',
    'MEDIAN',
    'ROUND',
    'ROUNDUP',
    'ROUNDDOWN',
    'CEILING',
    'FLOOR',
    'INT',
    'MOD',
    'ABS',
    'POWER',
    'SQRT',
    'EXP',
    'LN',
    'LOG',
    'LOG10',
  ],
  financial: ['PMT', 'PV', 'FV', 'NPV', 'IRR', 'RATE', 'NPER'],
  dynamicArray: [
    'TRANSPOSE',
    'UNIQUE',
    'SORT',
    'SORTBY',
    'FILTER',
    'SEQUENCE',
    'RANDARRAY',
    'VSTACK',
    'HSTACK',
    'TOROW',
    'TOCOL',
    'WRAPROWS',
    'WRAPCOLS',
    'CHOOSEROWS',
    'CHOOSECOLS',
    'TAKE',
    'DROP',
    'EXPAND',
    'LAMBDA',
    'LET',
    'MAP',
    'REDUCE',
    'SCAN',
    'BYROW',
    'BYCOL',
    'MAKEARRAY',
    'GROUPBY',
    'PIVOTBY',
    'PERCENTOF',
    'IMAGE',
  ],
};

export interface FxDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Optional spreadsheet-context seed for function arguments, e.g. the
   *  currently selected range when a seeded ribbon function opens. */
  getInitialArguments?: (functionName: string) => readonly string[] | null;
  /** Called with the assembled formula text (including leading '='). */
  onInsert: (formula: string) => void;
}

export interface FxDialogHandle {
  /** Open the dialog. Optional `seedName` pre-selects a function and jumps
   *  straight to the argument-input step. */
  open(seedName?: string): void;
  close(): void;
  /** Re-read i18n strings (e.g. after a locale switch). */
  refresh(): void;
  detach(): void;
}

/** Concise "spreadsheet-style" descriptions for the most common functions. The
 *  catalog itself (FUNCTION_SIGNATURES) doesn't carry descriptions — anything
 *  not listed here renders without a description blurb. Keep the list small;
 *  exhaustive coverage isn't a goal. */
export const FUNCTION_DESCRIPTIONS: Readonly<Record<string, { en: string; ja: string }>> = {
  SUM: { en: 'Adds its arguments.', ja: '引数の合計を返します。' },
  IF: {
    en: 'Returns one value when a condition is true and another when false.',
    ja: '条件が真のときと偽のときで異なる値を返します。',
  },
  VLOOKUP: {
    en: 'Looks up a value in the leftmost column of a table.',
    ja: '表の左端列で値を検索します。',
  },
  COUNT: { en: 'Counts numeric cells in the range.', ja: '範囲内の数値セルの個数を返します。' },
  COUNTA: {
    en: 'Counts non-empty cells in the range.',
    ja: '範囲内の空白でないセルの個数を返します。',
  },
  COUNTIF: {
    en: 'Counts cells in a range that match a condition.',
    ja: '条件を満たすセルの個数を返します。',
  },
  INDEX: {
    en: 'Returns a value at a given row/column in an array.',
    ja: '配列内の指定された行と列の値を返します。',
  },
  MATCH: {
    en: 'Returns the position of a value in an array.',
    ja: '配列内で一致する値の位置を返します。',
  },
  AVERAGE: { en: 'Returns the arithmetic mean.', ja: '引数の平均値を返します。' },
  MIN: { en: 'Returns the smallest argument.', ja: '引数の最小値を返します。' },
  MAX: { en: 'Returns the largest argument.', ja: '引数の最大値を返します。' },
  ROUND: {
    en: 'Rounds a number to a given precision.',
    ja: '数値を指定した桁数で四捨五入します。',
  },
  IFERROR: {
    en: 'Returns a fallback when the first argument is an error.',
    ja: '式がエラーの場合に代替値を返します。',
  },
  CONCAT: { en: 'Concatenates a list of texts.', ja: '複数の文字列を連結します。' },
  TEXT: {
    en: 'Formats a value as text using a format code.',
    ja: '書式コードに従って数値を文字列に整形します。',
  },
  LEFT: { en: 'Returns the left part of a string.', ja: '文字列の先頭から指定文字数を返します。' },
  RIGHT: {
    en: 'Returns the right part of a string.',
    ja: '文字列の末尾から指定文字数を返します。',
  },
  MID: {
    en: 'Returns characters from the middle of a string.',
    ja: '文字列の中間から指定文字数を返します。',
  },
  LEN: { en: 'Returns the length of a string.', ja: '文字列の文字数を返します。' },
  UPPER: { en: 'Converts a string to upper case.', ja: '文字列を大文字に変換します。' },
  LOWER: { en: 'Converts a string to lower case.', ja: '文字列を小文字に変換します。' },
  AND: { en: 'TRUE only when every argument is TRUE.', ja: 'すべての引数が真のとき真を返します。' },
  OR: { en: 'TRUE when any argument is TRUE.', ja: 'いずれかの引数が真のとき真を返します。' },
  NOT: { en: 'Inverts a boolean.', ja: '論理値を反転します。' },
  ISBLANK: { en: 'Tests whether a value is blank.', ja: '値が空白かどうかを返します。' },
  ISNUMBER: { en: 'Tests whether a value is numeric.', ja: '値が数値かどうかを返します。' },
  NOW: { en: 'Returns the current date and time.', ja: '現在の日付と時刻を返します。' },
  TODAY: { en: "Returns today's date.", ja: '今日の日付を返します。' },
  DATE: {
    en: 'Builds a date from year, month, day.',
    ja: '年・月・日からシリアル値を作成します。',
  },
};

/**
 * Spreadsheet-style "Function Arguments" modal. Two steps:
 *   1. Pick a function from a searchable list.
 *   2. Fill labeled inputs (one per declared arg in `FUNCTION_SIGNATURES`)
 *      with a live `= NAME(arg1, arg2, …)` preview and an Insert button.
 *
 * On confirm, calls `onInsert(formula)` with the assembled text and closes;
 * the caller is responsible for writing it into the active cell.
 */
export function attachFxDialog(deps: FxDialogDeps): FxDialogHandle {
  const { host, onInsert } = deps;
  let strings = deps.strings ?? defaultStrings;
  let t = strings.fxDialog;
  let locale: 'en' | 'ja' = detectLocale(strings);

  // ── Overlay + panel ─────────────────────────────────────────────────────
  const shell = createDialogShell({
    host,
    className: 'fc-fxdialog',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  // Reuse the shared format-dialog skin for header/footer/btn styling.
  const overlay = shell.overlay;
  overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-fxdialog__panel'],
    bodyClass: 'fc-fmtdlg__body fc-fxdialog__body',
  });

  // ── Step 1: function picker ─────────────────────────────────────────────
  const pickerWrap = document.createElement('div');
  pickerWrap.className = 'fc-fxdialog__picker';
  body.appendChild(pickerWrap);

  const categoryRow = document.createElement('label');
  categoryRow.className = 'fc-fxdialog__category-row';
  pickerWrap.appendChild(categoryRow);

  const categoryLabel = document.createElement('span');
  categoryLabel.textContent = t.categoryLabel;
  categoryRow.appendChild(categoryLabel);

  const categorySelect = createDialogSelect([], '', { className: 'fc-fxdialog__category' });
  categoryRow.appendChild(categorySelect);

  const searchInput = document.createElement('input');
  searchInput.type = 'text';
  searchInput.className = 'fc-fxdialog__search';
  searchInput.placeholder = t.searchPlaceholder;
  searchInput.setAttribute('aria-label', t.searchPlaceholder);
  searchInput.setAttribute('role', 'combobox');
  searchInput.setAttribute('aria-autocomplete', 'list');
  searchInput.setAttribute('aria-expanded', 'true');
  searchInput.autocomplete = 'off';
  searchInput.spellcheck = false;
  pickerWrap.appendChild(searchInput);

  const list = document.createElement('div');
  list.className = 'fc-fxdialog__list';
  list.setAttribute('role', 'listbox');
  list.setAttribute('aria-label', t.title);
  list.id = `fc-fxdialog-list-${Math.random().toString(36).slice(2, 8)}`;
  searchInput.setAttribute('aria-controls', list.id);
  pickerWrap.appendChild(list);

  const functionSummary = document.createElement('div');
  functionSummary.className = 'fc-fxdialog__function-summary';
  functionSummary.setAttribute('aria-live', 'polite');
  pickerWrap.appendChild(functionSummary);

  const functionSummaryName = document.createElement('div');
  functionSummaryName.className = 'fc-fxdialog__summary-name';
  functionSummary.appendChild(functionSummaryName);

  const functionSummaryDesc = document.createElement('div');
  functionSummaryDesc.className = 'fc-fxdialog__summary-desc';
  functionSummary.appendChild(functionSummaryDesc);

  // ── Step 2: argument inputs ─────────────────────────────────────────────
  const argsWrap = document.createElement('div');
  argsWrap.className = 'fc-fxdialog__args';
  argsWrap.hidden = true;
  body.appendChild(argsWrap);

  const argsHeader = document.createElement('div');
  argsHeader.className = 'fc-fxdialog__args-header';
  argsWrap.appendChild(argsHeader);

  const argsName = document.createElement('span');
  argsName.className = 'fc-fxdialog__args-name';
  argsHeader.appendChild(argsName);

  const argsDesc = document.createElement('div');
  argsDesc.className = 'fc-fxdialog__args-desc';
  argsWrap.appendChild(argsDesc);

  const argsFields = document.createElement('div');
  argsFields.className = 'fc-fxdialog__args-fields';
  argsWrap.appendChild(argsFields);

  const previewLabel = document.createElement('div');
  previewLabel.className = 'fc-fxdialog__preview-label';
  previewLabel.textContent = t.preview;
  argsWrap.appendChild(previewLabel);

  const preview = document.createElement('div');
  preview.className = 'fc-fxdialog__preview';
  argsWrap.appendChild(preview);

  // ── Footer ──────────────────────────────────────────────────────────────
  const backBtn = appendDialogButton(footer, { label: t.back });
  backBtn.style.marginRight = 'auto';
  backBtn.hidden = true;
  const { cancelBtn, okBtn: insertBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.insert,
  });
  projectDisabledState(insertBtn, true, t.insertRequiresFunction, {
    datasetKey: 'disabledReason',
    titlePrefix: t.insert,
  });

  // ── State ───────────────────────────────────────────────────────────────
  // Sorted catalog of all known function names; rebuilt once and reused.
  const allNames: readonly string[] = Object.keys(FUNCTION_SIGNATURES).slice().sort();
  let selectedCategory: FunctionCategory = 'all';
  let recentNames: string[] = [];
  let selectedName: string | null = null;
  let highlightIndex = 0;
  let argInputs: HTMLInputElement[] = [];

  const setInsertDisabled = (disabled: boolean, reason: string | null): void => {
    projectDisabledState(insertBtn, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: t.insert,
    });
  };

  const localizedDescription = (name: string): string => {
    const entry = FUNCTION_DESCRIPTIONS[name];
    if (!entry) return '';
    return locale === 'ja' ? entry.ja : entry.en;
  };

  const functionSyntax = (name: string): string =>
    `${name}(${(FUNCTION_SIGNATURES[name] ?? []).join(', ')})`;

  const updateFunctionSummary = (name: string | null): void => {
    if (!name) {
      functionSummary.hidden = true;
      functionSummaryName.textContent = '';
      functionSummaryDesc.textContent = '';
      return;
    }
    functionSummary.hidden = false;
    functionSummaryName.textContent = functionSyntax(name);
    functionSummaryDesc.textContent = localizedDescription(name);
  };

  const categoryOptions = (): Array<{ value: FunctionCategory; label: string }> => [
    { value: 'all', label: t.categoryAll },
    { value: 'recent', label: t.categoryRecent },
    { value: 'logical', label: t.categoryLogical },
    { value: 'lookup', label: t.categoryLookup },
    { value: 'text', label: t.categoryText },
    { value: 'datetime', label: t.categoryDateTime },
    { value: 'math', label: t.categoryMath },
    { value: 'financial', label: t.categoryFinancial },
    { value: 'dynamicArray', label: t.categoryDynamicArray },
  ];

  const renderCategoryOptions = (): void => {
    categorySelect.replaceChildren();
    appendDialogSelectOptions(categorySelect, categoryOptions());
    categorySelect.value = selectedCategory;
  };

  const categoryNames = (): string[] => {
    if (selectedCategory === 'all') return [...allNames];
    if (selectedCategory === 'recent')
      return recentNames.filter((name) => name in FUNCTION_SIGNATURES).slice().sort();
    return FUNCTION_CATEGORY_NAMES[selectedCategory]
      .filter((name) => name in FUNCTION_SIGNATURES)
      .slice()
      .sort();
  };

  const filteredNames = (): string[] => {
    const q = searchInput.value.trim().toUpperCase();
    const source = categoryNames();
    if (!q) return source;
    return source.filter((n) => n.includes(q));
  };

  const renderList = (): void => {
    list.replaceChildren();
    const names = filteredNames();
    if (names.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'fc-fxdialog__empty';
      empty.textContent = t.empty;
      list.appendChild(empty);
      highlightIndex = -1;
      searchInput.removeAttribute('aria-activedescendant');
      updateFunctionSummary(null);
      return;
    }
    if (highlightIndex < 0 || highlightIndex >= names.length) highlightIndex = 0;
    names.forEach((name, i) => {
      const item = document.createElement('div');
      item.className = 'fc-fxdialog__item';
      item.setAttribute('role', 'option');
      item.id = `${list.id}-option-${i}`;
      item.dataset.fxName = name;
      item.dataset.fxIndex = String(i);
      if (i === highlightIndex) {
        item.classList.add('fc-fxdialog__item--active');
        item.setAttribute('aria-selected', 'true');
      } else {
        item.setAttribute('aria-selected', 'false');
      }
      const nameEl = document.createElement('span');
      nameEl.className = 'fc-fxdialog__item-name';
      nameEl.textContent = name;
      item.appendChild(nameEl);
      const desc = localizedDescription(name);
      if (desc) {
        const descEl = document.createElement('span');
        descEl.className = 'fc-fxdialog__item-desc';
        descEl.textContent = desc;
        item.appendChild(descEl);
      }
      // No per-item listener — clicks bubble to the delegated handler on
      // `list`, registered once via shell.on() below. That keeps listener
      // count O(1) instead of O(n) and lets dispose() sweep them all.
      list.appendChild(item);
    });
    searchInput.setAttribute('aria-activedescendant', `${list.id}-option-${highlightIndex}`);
    updateFunctionSummary(names[highlightIndex] ?? null);
  };

  const assembleFormula = (): string => {
    if (!selectedName) return '';
    const args = argInputs.map((i) => i.value);
    // Drop trailing empties so `=SUM(1,,)` doesn't get assembled when only
    // the first slot was filled. Internal blanks are preserved as positional
    // placeholders.
    while (args.length > 0 && args[args.length - 1] === '') args.pop();
    return `=${selectedName}(${args.join(', ')})`;
  };

  const updatePreview = (): void => {
    preview.textContent = assembleFormula() || `=${selectedName ?? ''}()`;
    setInsertDisabled(!selectedName, selectedName ? null : t.insertRequiresFunction);
  };

  const choose = (name: string, initialArgs: readonly string[] = []): void => {
    selectedName = name;
    recentNames = [name, ...recentNames.filter((entry) => entry !== name)].slice(0, 12);
    pickerWrap.hidden = true;
    argsWrap.hidden = false;
    backBtn.hidden = false;
    setInsertDisabled(false, null);

    argsName.textContent = functionSyntax(name);
    argsDesc.textContent = localizedDescription(name);

    argsFields.replaceChildren();
    argInputs = [];
    const sig = FUNCTION_SIGNATURES[name] ?? [];
    // The trailing "..." marker in signatures is a hint — surface it as a
    // disabled label rather than a real input slot.
    const inputArgs = sig.filter((a) => a !== '...');
    inputArgs.forEach((arg) => {
      const row = document.createElement('label');
      row.className = 'fc-fmtdlg__row fc-fxdialog__arg-row';
      const labelEl = document.createElement('span');
      labelEl.textContent = arg;
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'fc-fxdialog__arg-input';
      input.autocomplete = 'off';
      input.spellcheck = false;
      // Tracked via shell.on so dispose() removes it; required because
      // step-2 inputs are dynamically rebuilt every choose() call.
      shell.on(input, 'input', updatePreview);
      row.append(labelEl, input);
      argsFields.appendChild(row);
      input.value = initialArgs[argInputs.length] ?? '';
      argInputs.push(input);
    });
    if (sig.includes('...')) {
      const note = document.createElement('div');
      note.className = 'fc-fxdialog__variadic-note';
      note.textContent = t.variadicHint;
      argsFields.appendChild(note);
    }
    updatePreview();
    requestAnimationFrame(() => {
      argInputs[0]?.focus();
    });
  };

  const goBackToPicker = (): void => {
    selectedName = null;
    pickerWrap.hidden = false;
    argsWrap.hidden = true;
    backBtn.hidden = true;
    setInsertDisabled(true, t.insertRequiresFunction);
    argInputs = [];
    requestAnimationFrame(() => searchInput.focus());
  };

  // ── Event handlers ──────────────────────────────────────────────────────
  const onSearchInput = (): void => {
    highlightIndex = 0;
    renderList();
  };

  const onCategoryChange = (): void => {
    selectedCategory = categorySelect.value as FunctionCategory;
    highlightIndex = 0;
    renderList();
  };

  const onSearchKey = (e: KeyboardEvent): void => {
    const names = filteredNames();
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (names.length === 0) return;
      highlightIndex = Math.min(highlightIndex + 1, names.length - 1);
      renderList();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (names.length === 0) return;
      highlightIndex = Math.max(highlightIndex - 1, 0);
      renderList();
    } else if (e.key === 'Home') {
      e.preventDefault();
      if (names.length === 0) return;
      highlightIndex = 0;
      renderList();
    } else if (e.key === 'End') {
      e.preventDefault();
      if (names.length === 0) return;
      highlightIndex = names.length - 1;
      renderList();
    } else if (e.key === 'Enter') {
      const target = names[highlightIndex];
      if (!target) return;
      e.preventDefault();
      e.stopPropagation();
      choose(target);
    }
  };

  const onInsertClick = (): void => {
    if (!selectedName) return;
    const formula = assembleFormula();
    onInsert(formula);
    api.close();
  };

  const onCancel = (): void => api.close();
  const onBack = (): void => goBackToPicker();

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
      return;
    }
    // Enter inside an arg input commits the assembled formula. The picker
    // step has its own Enter handler on the search input.
    if (e.key === 'Enter' && !argsWrap.hidden && !insertBtn.disabled) {
      e.preventDefault();
      onInsertClick();
    }
  };

  // Delegated picker click — fires for any rendered .fc-fxdialog__item via
  // bubble. Replaces the per-item listener that used to pile up on every
  // search-filter rerender and stayed unmatched in detach(). Listener count
  // is now O(1) regardless of how many functions are visible.
  const onListClick = (e: Event): void => {
    const target = (e.target as HTMLElement | null)?.closest<HTMLElement>('.fc-fxdialog__item');
    if (!target?.dataset.fxName) return;
    const idx = Number.parseInt(target.dataset.fxIndex ?? '-1', 10);
    if (Number.isFinite(idx) && idx >= 0) highlightIndex = idx;
    choose(target.dataset.fxName);
  };

  shell.on(list, 'click', onListClick);
  shell.on(categorySelect, 'change', onCategoryChange);
  shell.on(searchInput, 'input', onSearchInput);
  shell.on(searchInput, 'keydown', onSearchKey as EventListener);
  shell.on(insertBtn, 'click', onInsertClick);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(backBtn, 'click', onBack);
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  const refreshLabels = (): void => {
    t = strings.fxDialog;
    shell.setAriaLabel(t.title);
    header.textContent = t.title;
    categoryLabel.textContent = t.categoryLabel;
    renderCategoryOptions();
    searchInput.placeholder = t.searchPlaceholder;
    searchInput.setAttribute('aria-label', t.searchPlaceholder);
    list.setAttribute('aria-label', t.title);
    previewLabel.textContent = t.preview;
    backBtn.textContent = t.back;
    cancelBtn.textContent = t.cancel;
    insertBtn.textContent = t.insert;
    setInsertDisabled(insertBtn.disabled, insertBtn.disabled ? t.insertRequiresFunction : null);
    if (selectedName) {
      argsDesc.textContent = localizedDescription(selectedName);
      const noteEl = argsFields.querySelector<HTMLElement>('.fc-fxdialog__variadic-note');
      if (noteEl) noteEl.textContent = t.variadicHint;
    }
    renderList();
  };

  const api: FxDialogHandle = {
    open(seedName?: string): void {
      searchInput.value = '';
      highlightIndex = 0;
      argInputs = [];
      renderCategoryOptions();
      const seed = seedName ? seedName.toUpperCase() : null;
      if (seed && FUNCTION_SIGNATURES[seed]) {
        // Skip the picker — jump straight to argument entry.
        choose(seed, deps.getInitialArguments?.(seed) ?? []);
      } else {
        selectedName = null;
        pickerWrap.hidden = false;
        argsWrap.hidden = true;
        backBtn.hidden = true;
        setInsertDisabled(true, t.insertRequiresFunction);
        renderList();
      }
      shell.open();
      requestAnimationFrame(() => {
        if (argsWrap.hidden) searchInput.focus();
        else argInputs[0]?.focus();
      });
    },
    close(): void {
      shell.close();
      host.focus();
    },
    refresh(): void {
      // Re-snapshot strings from the original deps reference. The caller
      // ferries the latest dictionary in via setStrings-style updates.
      strings = deps.strings ?? defaultStrings;
      locale = detectLocale(strings);
      refreshLabels();
    },
    detach(): void {
      shell.dispose();
    },
  };

  // First paint of the picker so a synchronous open() can render without a
  // microtask in tests.
  renderCategoryOptions();
  renderList();

  return api;
}
