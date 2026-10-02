import {
  buildFunctionCatalog,
  type FunctionCatalogEntry,
  type FunctionCatalogReader,
  type FunctionCatalogSnapshot,
  type FunctionCategory,
  supportedFunctionNames,
} from '../commands/function-categories.js';
import {
  getRecentFunctions,
  recordRecentFunction,
  subscribeRecentFunctions,
} from '../commands/function-history.js';
import { FUNCTION_SIGNATURES } from '../commands/refs.js';
import { defaultStrings, en as enStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  createDialogButton,
  createDialogShell,
} from './dialog-shell.js';

/** Heuristic locale detector — we don't get a `Locale` flag through deps so
 *  we sniff the active dictionary's title against the canonical English one.
 *  Cheap, and correct for the two built-in locales. Custom locales fall back
 *  to English descriptions, which is the standard desktop spreadsheets behaviour for
 *  unsupported tongues. */
const detectLocale = (s: Strings): 'en' | 'ja' =>
  s.fxDialog.title === enStrings.fxDialog.title ? 'en' : 'ja';

/** Keep the picker-facing type exported from its historic module for hosts
 * that import it from `interact/fx-dialog`. */
export type { FunctionCategory } from '../commands/function-categories.js';

export interface FxDialogOpenOptions {
  /** Start on a picker category. Ignored when `seedName` names a function. */
  category?: FunctionCategory;
}

export interface FxDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Optional spreadsheet-context seed for function arguments, e.g. the
   *  currently selected range when a seeded ribbon function opens. */
  getInitialArguments?: (functionName: string) => readonly string[] | null;
  /** Read the live recognized function catalog at each open/refresh. */
  getWb?: () => FunctionCatalogReader | null;
  /** Current host locale, used for engine-provided function metadata. */
  getLocale?: () => string;
  /** Called with the assembled formula text (including leading '='). */
  onInsert: (formula: string) => void | boolean;
}

export interface FxDialogHandle {
  /** Open the dialog. Optional `seedName` pre-selects a function and jumps
   *  straight to the argument-input step. */
  open(seedName?: string, options?: FxDialogOpenOptions): void;
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
  const localeOrdinal = (): 0 | 1 =>
    (deps.getLocale?.().toLowerCase().startsWith('ja') ?? locale === 'ja') ? 1 : 0;

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
  // The workbook catalog is intentionally lazy: a workbook can be swapped
  // after attachment, and locale metadata can change between opens.
  let catalog: FunctionCatalogSnapshot = buildFunctionCatalog(
    deps.getWb?.() ?? null,
    localeOrdinal(),
  );
  let selectedCategory: FunctionCategory = 'all';
  let selectedName: string | null = null;
  let highlightIndex = 0;
  let argInputs: HTMLInputElement[] = [];
  let argCount = 0;
  let selectedEntry: FunctionCatalogEntry | null = null;

  const refreshCatalog = (): void => {
    locale = deps.getLocale?.().toLowerCase().startsWith('ja') ? 'ja' : detectLocale(strings);
    catalog = buildFunctionCatalog(deps.getWb?.() ?? null, localeOrdinal());
  };

  const setInsertDisabled = (disabled: boolean, reason: string | null): void => {
    projectDisabledState(insertBtn, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: t.insert,
    });
  };

  const catalogEntry = (name: string): FunctionCatalogEntry | null =>
    catalog.entries.get(name) ?? null;

  const localizedDescription = (name: string): string => {
    const metadataDescription = catalogEntry(name)?.description;
    if (metadataDescription !== undefined) return metadataDescription;
    const entry = FUNCTION_DESCRIPTIONS[name];
    if (!entry) return '';
    return locale === 'ja' ? entry.ja : entry.en;
  };

  const functionSyntax = (name: string): string => {
    const entry = catalogEntry(name);
    if (entry?.signatureTemplate) return entry.signatureTemplate;
    const displayName = entry?.displayName ?? name;
    return `${displayName}(${(entry?.argumentLabels ?? FUNCTION_SIGNATURES[name] ?? []).join(', ')})`;
  };

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

  const categoryOptions = (): Array<{ value: FunctionCategory; label: string }> => {
    // The base dictionaries predate the native Excel family expansion. Keep
    // those dictionaries source-compatible and use stable labels until a
    // host supplies localized overrides for the newer families.
    const extended = t as typeof t &
      Partial<
        Record<
          | 'categoryStatistical'
          | 'categoryEngineering'
          | 'categoryInformation'
          | 'categoryDatabase'
          | 'categoryCompatibility'
          | 'categoryCube'
          | 'categoryWeb',
          string
        >
      >;
    const label = (key: keyof typeof extended, fallback: string): string =>
      extended[key] ?? fallback;
    const options: Array<{ value: FunctionCategory; label: string }> = [
      { value: 'all', label: t.categoryAll },
      { value: 'recent', label: t.categoryRecent },
      { value: 'financial', label: t.categoryFinancial },
      { value: 'logical', label: t.categoryLogical },
      { value: 'text', label: t.categoryText },
      { value: 'datetime', label: t.categoryDateTime },
      { value: 'lookup', label: t.categoryLookup },
      { value: 'math', label: t.categoryMath },
      {
        value: 'statistical',
        label: label('categoryStatistical', locale === 'ja' ? '統計' : 'Statistical'),
      },
      { value: 'engineering', label: label('categoryEngineering', 'Engineering') },
      { value: 'information', label: label('categoryInformation', 'Information') },
      { value: 'database', label: label('categoryDatabase', 'Database') },
      { value: 'compatibility', label: label('categoryCompatibility', 'Compatibility') },
      { value: 'cube', label: label('categoryCube', 'Cube') },
      { value: 'web', label: label('categoryWeb', 'Web') },
      { value: 'dynamicArray', label: t.categoryDynamicArray },
    ];
    return options.filter((option) => {
      if (option.value === 'all' || option.value === 'recent') return true;
      return supportedFunctionNames(option.value, catalog.knownNames).length > 0;
    });
  };

  const renderCategoryOptions = (): void => {
    categorySelect.replaceChildren();
    appendDialogSelectOptions(categorySelect, categoryOptions());
    categorySelect.value = selectedCategory;
  };

  const categoryNames = (): string[] => {
    if (selectedCategory === 'all') return [...catalog.names];
    if (selectedCategory === 'recent')
      return [...getRecentFunctions(deps.store, catalog.knownNames)];
    return supportedFunctionNames(selectedCategory, catalog.knownNames);
  };

  const filteredNames = (): string[] => {
    const q = searchInput.value.trim().toUpperCase();
    const source = categoryNames();
    if (!q) return source;
    return source.filter((name) => {
      const entry = catalogEntry(name);
      return [name, entry?.displayName, entry?.signatureTemplate, entry?.description]
        .filter((value): value is string => value !== undefined)
        .some((value) => value.toUpperCase().includes(q));
    });
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
      nameEl.textContent = catalogEntry(name)?.displayName ?? name;
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

  const unsubscribeRecentFunctions = subscribeRecentFunctions(deps.store, () => {
    if (selectedCategory !== 'recent') return;
    highlightIndex = 0;
    renderList();
  });

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

  const argumentCountFor = (entry: FunctionCatalogEntry | null, initialCount: number): number => {
    const labels = entry?.argumentLabels ?? [];
    const minArity = entry?.minArity ?? labels.length;
    const maxArity = entry?.maxArity ?? null;
    const optionalSeed = minArity === 0 && maxArity !== 0 ? 1 : 0;
    const requested = Math.max(labels.length, minArity, optionalSeed, initialCount);
    return maxArity === null ? requested : Math.min(requested, maxArity);
  };

  const argumentLabelFor = (entry: FunctionCatalogEntry | null, index: number): string => {
    const label = entry?.argumentLabels[index];
    if (label) return label;
    const minArity = entry?.minArity ?? 0;
    const prefix =
      index < minArity
        ? (t.argumentLabel ?? 'Argument')
        : (t.optionalArgumentLabel ?? 'Optional argument');
    return `${prefix} ${index + 1}`;
  };

  const renderArgumentFields = (initialArgs: readonly string[] = []): void => {
    const previousValues = argInputs.map((input) => input.value);
    argsFields.replaceChildren();
    argInputs = [];
    const entry = selectedEntry;
    const minArity = entry?.minArity ?? 0;
    const maxArity = entry?.maxArity ?? null;
    for (let index = 0; index < argCount; index += 1) {
      const row = document.createElement('label');
      row.className = 'fc-fmtdlg__row fc-fxdialog__arg-row';
      const labelEl = document.createElement('span');
      labelEl.textContent = argumentLabelFor(entry, index);
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'fc-fxdialog__arg-input';
      input.autocomplete = 'off';
      input.spellcheck = false;
      row.append(labelEl, input);
      argsFields.appendChild(row);
      input.value = initialArgs[index] ?? previousValues[index] ?? '';
      argInputs.push(input);
    }
    if (maxArity === null) {
      const note = document.createElement('div');
      note.className = 'fc-fxdialog__variadic-note';
      note.textContent = t.variadicHint;
      argsFields.appendChild(note);
    }
    if (maxArity === null || argCount < maxArity) {
      const add = createDialogButton({
        label: t.addArgument ?? 'Add argument',
        baseClass: 'fc-fxdialog__arg-action',
      });
      add.dataset.fxAction = 'add-argument';
      argsFields.appendChild(add);
    }
    if (argCount > minArity) {
      const remove = createDialogButton({
        label: t.removeArgument ?? 'Remove argument',
        baseClass: 'fc-fxdialog__arg-action',
      });
      remove.dataset.fxAction = 'remove-argument';
      argsFields.appendChild(remove);
    }
  };

  const choose = (name: string, initialArgs: readonly string[] = []): void => {
    selectedName = name;
    selectedEntry = catalogEntry(name);
    pickerWrap.hidden = true;
    argsWrap.hidden = false;
    backBtn.hidden = false;
    setInsertDisabled(false, null);

    argsName.textContent = functionSyntax(name);
    argsDesc.textContent = localizedDescription(name);

    argCount = argumentCountFor(selectedEntry, initialArgs.length);
    renderArgumentFields(initialArgs);
    updatePreview();
    if (shell.isOpen()) argInputs[0]?.focus();
  };

  const goBackToPicker = (): void => {
    selectedName = null;
    pickerWrap.hidden = false;
    argsWrap.hidden = true;
    backBtn.hidden = true;
    setInsertDisabled(true, t.insertRequiresFunction);
    argInputs = [];
    argCount = 0;
    selectedEntry = null;
    if (shell.isOpen()) searchInput.focus();
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
    let accepted: void | boolean;
    try {
      accepted = onInsert(formula);
    } catch {
      insertBtn.focus();
      return;
    }
    if (accepted === false) {
      insertBtn.focus();
      return;
    }
    recordRecentFunction(deps.store, selectedName, catalog.knownNames);
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

  const onArgsInput = (): void => updatePreview();

  const onArgsClick = (e: Event): void => {
    const target = (e.target as HTMLElement | null)?.closest<HTMLButtonElement>('[data-fx-action]');
    if (!target) return;
    const entry = selectedEntry;
    const minArity = entry?.minArity ?? 0;
    const maxArity = entry?.maxArity ?? null;
    if (target.dataset.fxAction === 'add-argument') {
      if (maxArity !== null && argCount >= maxArity) return;
      argCount += 1;
      renderArgumentFields();
      argInputs.at(-1)?.focus();
      updatePreview();
    } else if (target.dataset.fxAction === 'remove-argument') {
      if (argCount <= minArity) return;
      argCount -= 1;
      renderArgumentFields();
      argInputs.at(-1)?.focus();
      updatePreview();
    }
  };

  shell.on(list, 'click', onListClick);
  shell.on(argsFields, 'input', onArgsInput as EventListener, true);
  shell.on(argsFields, 'click', onArgsClick);
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
    refreshCatalog();
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
      selectedEntry = catalogEntry(selectedName);
      argsName.textContent = functionSyntax(selectedName);
      argsDesc.textContent = localizedDescription(selectedName);
      renderArgumentFields();
    }
    renderList();
  };

  const api: FxDialogHandle = {
    open(seedName?: string, options?: FxDialogOpenOptions): void {
      refreshCatalog();
      searchInput.value = '';
      highlightIndex = 0;
      argInputs = [];
      selectedCategory = options?.category ?? 'all';
      renderCategoryOptions();
      // Repaint the hidden picker too. Seeded opens skip directly to the
      // argument step, but the picker remains mounted and must reflect a
      // workbook or locale change before the next back/open cycle.
      renderList();
      const seed = seedName ? seedName.toUpperCase() : null;
      if (seed && catalog.knownNames.has(seed)) {
        // Skip the picker — jump straight to argument entry.
        choose(seed, deps.getInitialArguments?.(seed) ?? []);
      } else {
        selectedName = null;
        selectedEntry = null;
        pickerWrap.hidden = false;
        argsWrap.hidden = true;
        backBtn.hidden = true;
        setInsertDisabled(true, t.insertRequiresFunction);
      }
      shell.open();
      if (argsWrap.hidden) searchInput.focus();
      else argInputs[0]?.focus();
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
      unsubscribeRecentFunctions();
      shell.dispose();
    },
  };

  // First paint of the picker so a synchronous open() can render without a
  // microtask in tests.
  renderCategoryOptions();
  renderList();

  return api;
}
