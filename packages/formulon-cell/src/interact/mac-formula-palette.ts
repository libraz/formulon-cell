import {
  buildFunctionCatalog,
  type CatalogFunctionCategory,
  type FunctionCatalogEntry,
  type FunctionCatalogReader,
  type FunctionCatalogSnapshot,
  type FunctionCategory,
  functionSyntax,
  isFunctionUnavailableForInsertion,
  supportedFunctionNames,
} from '../commands/function-categories.js';
import {
  getRecentFunctions,
  recordRecentFunction,
  subscribeRecentFunctions,
} from '../commands/function-history.js';
import type { Addr } from '../engine/types.js';
import { formatCell, fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { ExternalFormulaDraftHandle, FormulaBarController } from '../mount/formula-bar.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import type { FormulaEditLease, FormulaEditLeaseContext } from './formula-edit-lease.js';
import {
  FUNCTION_DESCRIPTIONS,
  type FxDialogHandle,
  type FxDialogOpenOptions,
} from './fx-dialog.js';
import { assembledFormula, canonicalName, parseOuterCall } from './mac-formula-call.js';
import { makeButton, makeIconButton } from './mac-formula-palette-buttons.js';
import type { RangeInsertTarget } from './pointer.js';

type MacPaletteStrings = Strings['fxDialog']['macPalette'];

export interface MacFormulaArgumentHelp {
  label?: string;
  description?: string;
  url?: string;
}

export interface MacFormulaPaletteDeps {
  host: HTMLElement;
  dock: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle;
  getLocale: () => string;
  getStrings: () => Strings;
  getAnchor: () => Addr;
  getInitialArguments?: (name: string) => readonly string[] | null;
  beginDraft: FormulaBarController['beginExternalDraft'];
  /** Suspend an in-progress cell or formula-bar edit so the palette can take it over. */
  suspendActiveEdit?: (context: FormulaEditLeaseContext) => FormulaEditLease | null;
  projectMirror: (anchor: Addr, raw: string | null) => void;
  getArgumentHelp?: (name: string, index: number, locale: string) => MacFormulaArgumentHelp | null;
  getLocalizedArgumentHelp?: (
    name: string,
    index: number,
    locale: string,
  ) => MacFormulaArgumentHelp | null;
}

export interface MacFormulaPaletteHandle extends FxDialogHandle {
  isOpen(): boolean;
  rangeInsertTarget(): RangeInsertTarget | null;
  setStrings(next: Strings): void;
}

type PaletteMode = 'closed' | 'picker' | 'arguments-editing' | 'arguments-committed';

type CategoryLabelKey =
  | 'categoryLogical'
  | 'categoryLookup'
  | 'categoryText'
  | 'categoryDateTime'
  | 'categoryMath'
  | 'categoryFinancial'
  | 'categoryDynamicArray'
  | 'categoryStatistical'
  | 'categoryEngineering'
  | 'categoryInformation'
  | 'categoryDatabase'
  | 'categoryCompatibility'
  | 'categoryCube'
  | 'categoryWeb';

interface CommitSnapshot {
  name: string;
  args: string[];
  raw: string;
  result: string;
}

interface DraftBinding {
  handle: ExternalFormulaDraftHandle;
  unsubscribe: () => void;
  leased: boolean;
}

interface PickerRefs {
  search: HTMLInputElement;
  sections: HTMLElement;
  insert: HTMLButtonElement;
}

interface ArgumentsRefs {
  fields: HTMLElement;
  preview: HTMLElement;
  done: HTMLButtonElement;
}

const paletteStrings = (strings: Strings): MacPaletteStrings => strings.fxDialog.macPalette;

const localeOrdinal = (locale: string): 0 | 1 =>
  locale.trim().toLowerCase().startsWith('ja') ? 1 : 0;

const sameAvailability = (a: FunctionCatalogEntry, b: FunctionCatalogEntry): boolean =>
  Object.is(a.availability, b.availability);

const connectedElement = (value: Element | null): HTMLElement | null =>
  value instanceof HTMLElement && value.isConnected ? value : null;

export function attachMacFormulaPalette(deps: MacFormulaPaletteDeps): MacFormulaPaletteHandle {
  const root = document.createElement('aside');
  root.className = 'fc-mac-formula-palette';
  root.setAttribute('role', 'complementary');
  root.style.inlineSize = '300px';
  root.style.flexBasis = '300px';
  root.hidden = true;
  deps.dock.appendChild(root);
  deps.dock.hidden = true;

  let strings = deps.getStrings();
  let labels = paletteStrings(strings);
  let mode: PaletteMode = 'closed';
  let detached = false;
  let opener: HTMLElement | null = null;
  let restoredFocusTarget: HTMLElement | null = null;
  let anchor: Addr | null = null;
  let sessionWb: WorkbookHandle | null = null;
  let catalog = buildFunctionCatalog(null, localeOrdinal(deps.getLocale()));
  let pickerCategory: FunctionCategory = 'all';
  let pickerSelectionName: string | null = null;
  let pickerSelectionEntry: FunctionCatalogEntry | null = null;
  let selectedName: string | null = null;
  let selectedEntry: FunctionCatalogEntry | null = null;
  let args: string[] = [];
  let synchronized = true;
  let explicitArgumentCount: number | null = null;
  let preserveExplicitArgumentCount = false;
  let writingDraftRaw = false;
  let focusedArgument = 0;
  let currentRaw = '';
  let searchQuery = '';
  let guardMessage = '';
  let postCommit: CommitSnapshot | null = null;
  let activeDraft: DraftBinding | null = null;
  let pickerRefs: PickerRefs | null = null;
  let argumentsRefs: ArgumentsRefs | null = null;
  let recentUnsubscribe: (() => void) | null = null;
  let closing = false;
  let guarding = false;
  let committing = false;

  const leaseContext: FormulaEditLeaseContext = {
    getLocale: () => deps.getLocale(),
    contextCurrent: () => !detached && mode !== 'closed' && safeWorkbook() === sessionWb,
  };

  const currentAnchor = (): Addr => anchor ?? { ...deps.getAnchor() };

  const safeWorkbook = (): WorkbookHandle | null => {
    try {
      return deps.getWb();
    } catch {
      return null;
    }
  };

  const readCatalog = (): FunctionCatalogSnapshot => {
    let reader: FunctionCatalogReader | null = null;
    try {
      reader = deps.getWb();
      return buildFunctionCatalog(reader, localeOrdinal(deps.getLocale()));
    } catch {
      return buildFunctionCatalog(null, localeOrdinal(deps.getLocale()));
    }
  };

  const textForEntry = (entry: FunctionCatalogEntry): string =>
    entry.description ??
    (localeOrdinal(deps.getLocale()) === 1
      ? FUNCTION_DESCRIPTIONS[entry.canonicalName]?.ja
      : FUNCTION_DESCRIPTIONS[entry.canonicalName]?.en) ??
    '';

  const argumentHelp = (entry: FunctionCatalogEntry, index: number): MacFormulaArgumentHelp => {
    const provider = deps.getArgumentHelp ?? deps.getLocalizedArgumentHelp;
    const provided = provider?.(entry.canonicalName, index, deps.getLocale()) ?? null;
    const fallbackLabel = entry.argumentLabels[index] ?? `${labels.argument} ${index + 1}`;
    return {
      label: provided?.label ?? fallbackLabel.replace(/^\[|\]$/g, ''),
      description: provided?.description,
      url: provided?.url,
    };
  };

  const notifyMirror = (raw: string | null): void => {
    if (!anchor) return;
    try {
      deps.projectMirror({ ...anchor }, raw);
    } catch {
      // A mirror is presentation-only and cannot interrupt draft ownership.
    }
  };

  const clearRoot = (): void => {
    while (root.firstChild) root.removeChild(root.firstChild);
  };

  const selectedDescription = (): string => {
    if (!selectedEntry) return '';
    return textForEntry(selectedEntry) ?? '';
  };

  const selectedSyntax = (): string => (selectedEntry ? functionSyntax(selectedEntry) : '');

  const categoryTitle = (category: CatalogFunctionCategory): string => {
    const fx = strings.fxDialog;
    const keyByCategory: Record<CatalogFunctionCategory, CategoryLabelKey> = {
      logical: 'categoryLogical',
      lookup: 'categoryLookup',
      text: 'categoryText',
      datetime: 'categoryDateTime',
      math: 'categoryMath',
      financial: 'categoryFinancial',
      dynamicArray: 'categoryDynamicArray',
      statistical: 'categoryStatistical',
      engineering: 'categoryEngineering',
      information: 'categoryInformation',
      database: 'categoryDatabase',
      compatibility: 'categoryCompatibility',
      cube: 'categoryCube',
      web: 'categoryWeb',
    };
    const key = keyByCategory[category];
    return fx[key];
  };

  const reconcilePickerSelection = (): void => {
    if (!pickerSelectionName) {
      pickerSelectionEntry = null;
      return;
    }
    const entry = catalog.entries.get(pickerSelectionName) ?? null;
    if (
      !catalog.knownNames.has(pickerSelectionName) ||
      !entry ||
      isFunctionUnavailableForInsertion(entry.availability)
    ) {
      pickerSelectionName = null;
      pickerSelectionEntry = null;
      return;
    }
    pickerSelectionEntry = entry;
  };

  const updateDataset = (): void => {
    root.dataset.state = mode;
    root.dataset.rawSynchronized = synchronized ? 'true' : 'false';
    root.hidden = mode === 'closed';
    deps.dock.hidden = mode === 'closed' || detached;
    root.setAttribute('aria-label', labels.title);
  };

  const updateFieldValues = (): void => {
    const fields = argumentsRefs?.fields.querySelectorAll<HTMLInputElement>(
      'input[data-argument-index]',
    );
    fields?.forEach((field) => {
      const index = Number(field.dataset.argumentIndex);
      field.value = args[index] ?? '';
    });
  };

  const renderPreview = (): void => {
    const preview = argumentsRefs?.preview;
    if (!preview || !selectedEntry) return;
    preview.textContent = '';
    const pendingRequired =
      synchronized && args.slice(0, selectedEntry.minArity).some((value) => value.trim() === '');
    if (pendingRequired || !currentRaw.trim().startsWith('=')) {
      preview.textContent = labels.pending;
      return;
    }
    try {
      const workbook = safeWorkbook();
      if (!workbook) {
        preview.textContent = labels.pending;
        return;
      }
      const result = workbook.evaluateFormulaText(currentAnchor(), currentRaw);
      if (!result.status.ok) {
        preview.textContent = labels.pending;
        return;
      }
      preview.textContent = formatCell(fromEngineValue(result.value), deps.getLocale());
    } catch {
      preview.textContent = labels.pending;
    }
  };

  const updatePickerList = (): void => {
    const sections = pickerRefs?.sections;
    if (!sections) return;
    sections.replaceChildren();
    const query = searchQuery.trim().toUpperCase();
    const namesFor = (names: readonly string[]): string[] =>
      names.filter((name) => !query || name.includes(query));

    const appendSection = (key: string, title: string, names: readonly string[]): void => {
      const section = document.createElement('section');
      section.className = 'fc-mac-formula-palette__section';
      section.dataset.section = key;
      const heading = document.createElement('h3');
      heading.textContent = title;
      section.appendChild(heading);
      const list = document.createElement('div');
      list.className = 'fc-mac-formula-palette__function-list';
      list.setAttribute('role', 'listbox');
      for (const name of namesFor(names)) {
        const entry = catalog.entries.get(name);
        if (!entry) continue;
        const row = document.createElement('div');
        row.className = 'fc-mac-formula-palette__function';
        row.dataset.functionName = entry.canonicalName;
        row.setAttribute('role', 'option');
        row.tabIndex = 0;
        row.setAttribute(
          'aria-selected',
          String((pickerSelectionName ?? selectedName) === entry.canonicalName),
        );
        const unavailable = isFunctionUnavailableForInsertion(entry.availability);
        if (unavailable) {
          row.classList.add('is-unavailable');
        }
        projectDisabledState(row, unavailable, unavailable ? labels.unavailable : null);
        const nameSpan = document.createElement('span');
        nameSpan.className = 'fc-mac-formula-palette__function-name';
        nameSpan.textContent = entry.displayName;
        row.appendChild(nameSpan);
        const description = textForEntry(entry);
        if (description) {
          const descriptionSpan = document.createElement('span');
          descriptionSpan.className = 'fc-mac-formula-palette__function-description';
          descriptionSpan.textContent = description;
          row.appendChild(descriptionSpan);
        }
        row.addEventListener('click', () => {
          if (detached || mode !== 'picker' || !row.isConnected) return;
          const currentEntry = catalog.entries.get(entry.canonicalName) ?? null;
          if (
            !catalog.knownNames.has(entry.canonicalName) ||
            !currentEntry ||
            isFunctionUnavailableForInsertion(currentEntry.availability)
          ) {
            reconcilePickerSelection();
            render();
            return;
          }
          pickerSelectionName = currentEntry.canonicalName;
          pickerSelectionEntry = currentEntry;
          guardMessage = '';
          render();
        });
        row.addEventListener('keydown', (event) => {
          if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            row.click();
          }
        });
        list.appendChild(row);
      }
      if (!list.childElementCount) {
        const empty = document.createElement('p');
        empty.className = 'fc-mac-formula-palette__empty';
        empty.textContent = labels.empty;
        list.appendChild(empty);
      }
      section.appendChild(list);
      sections.appendChild(section);
    };

    if (pickerCategory === 'all') {
      appendSection('recent', labels.recent, getRecentFunctions(deps.store, catalog.knownNames));
      appendSection('all', labels.all, catalog.names);
    } else if (pickerCategory === 'recent') {
      appendSection('recent', labels.recent, getRecentFunctions(deps.store, catalog.knownNames));
    } else {
      appendSection(
        pickerCategory,
        categoryTitle(pickerCategory),
        supportedFunctionNames(pickerCategory, catalog.knownNames),
      );
    }
  };

  const updatePickerSummary = (summary: HTMLElement): void => {
    summary.replaceChildren();
    const entry = pickerSelectionEntry ?? selectedEntry;
    if (!entry) return;
    const name = document.createElement('strong');
    name.textContent = entry.displayName;
    summary.appendChild(name);
    const description = textForEntry(entry);
    if (description) {
      const text = document.createElement('span');
      text.textContent = description;
      summary.appendChild(text);
    }
    const syntax = document.createElement('code');
    syntax.textContent = functionSyntax(entry);
    summary.appendChild(syntax);
  };

  const selectedUnavailable = (): boolean =>
    selectedEntry !== null && isFunctionUnavailableForInsertion(selectedEntry.availability);

  const showGuardedPicker = (message: string): void => {
    guarding = true;
    if (activeDraft) activeDraft.handle.cancel();
    guarding = false;
    const live = safeWorkbook();
    sessionWb = live;
    catalog = readCatalog();
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
    pickerCategory = 'all';
    pickerSelectionName = null;
    pickerSelectionEntry = null;
    mode = 'picker';
    guardMessage = message;
    currentRaw = '';
    explicitArgumentCount = null;
    preserveExplicitArgumentCount = false;
    synchronized = true;
    notifyMirror(null);
    render();
  };

  const liveEntryGuard = (): FunctionCatalogEntry | null => {
    if (!selectedName) return null;
    const live = safeWorkbook();
    const nextCatalog = readCatalog();
    const nextEntry = nextCatalog.entries.get(selectedName) ?? null;
    const mismatch =
      live === null ||
      live !== sessionWb ||
      !nextCatalog.knownNames.has(selectedName) ||
      nextEntry === null ||
      selectedEntry === null ||
      !sameAvailability(selectedEntry, nextEntry);
    catalog = nextCatalog;
    reconcilePickerSelection();
    if (mismatch) {
      showGuardedPicker(labels.unavailable);
      return null;
    }
    sessionWb = live;
    selectedEntry = nextEntry;
    return nextEntry;
  };

  const setDraftRaw = (raw: string, preserveExplicit = preserveExplicitArgumentCount): void => {
    if (!activeDraft) return;
    preserveExplicitArgumentCount = preserveExplicit;
    writingDraftRaw = true;
    try {
      activeDraft.handle.setValue(raw);
    } finally {
      writingDraftRaw = false;
    }
    currentRaw = raw;
  };

  const assembledCurrentFormula = (name: string): string => {
    if (!preserveExplicitArgumentCount || explicitArgumentCount === null)
      return assembledFormula(name, args);
    return `=${name}(${args.slice(0, explicitArgumentCount).join(',')})`;
  };

  const handleRawUpdate = (binding: DraftBinding, raw: string): void => {
    if (activeDraft !== binding) return;
    currentRaw = raw;
    notifyMirror(raw);
    if (selectedName) {
      const parsed = parseOuterCall(raw, selectedName);
      if (parsed) {
        const changedCount = args.length !== parsed.args.length;
        args = parsed.args;
        explicitArgumentCount = parsed.args.length;
        if (!writingDraftRaw) preserveExplicitArgumentCount = true;
        synchronized = true;
        if (changedCount && mode !== 'picker') render();
        updateFieldValues();
      } else {
        synchronized = false;
        explicitArgumentCount = null;
        preserveExplicitArgumentCount = false;
      }
    } else {
      synchronized = true;
    }
    updateDataset();
    renderPreview();
  };

  const finishDraft = (binding: DraftBinding, outcome: 'committed' | 'cancelled'): void => {
    if (activeDraft !== binding) return;
    const committedName = selectedName;
    const committedRaw = currentRaw;
    const committedArgs = [...args];
    const committedResult = argumentsRefs?.preview.textContent ?? labels.pending;
    const committedCount = explicitArgumentCount;
    const committedPreservesCount = preserveExplicitArgumentCount;
    const parsed =
      outcome === 'committed' && committedName ? parseOuterCall(committedRaw, committedName) : null;
    const assembled =
      parsed && committedName
        ? committedPreservesCount && committedCount !== null
          ? `=${committedName}(${parsed.args.slice(0, committedCount).join(',')})`
          : assembledFormula(committedName, parsed.args)
        : null;
    const shouldRecord = parsed !== null && assembled === committedRaw.trim();
    activeDraft = null;
    binding.unsubscribe();
    if (outcome === 'committed' && committedName) {
      postCommit = {
        name: committedName,
        args: committedArgs,
        raw: committedRaw,
        result: committedResult,
      };
      mode = 'arguments-committed';
      synchronized = parsed !== null;
      explicitArgumentCount = parsed?.args.length ?? null;
      preserveExplicitArgumentCount = parsed !== null;
      guardMessage = '';
      if (shouldRecord) recordRecentFunction(deps.store, committedName, catalog.knownNames);
      notifyMirror(null);
      if (detached || closing || mode !== 'arguments-committed') return;
      render();
      return;
    }
    notifyMirror(null);
    if (outcome === 'cancelled' && !closing && !guarding && !committing && mode !== 'closed') {
      mode = 'picker';
      selectedName = null;
      selectedEntry = null;
      args = [];
      explicitArgumentCount = null;
      preserveExplicitArgumentCount = false;
      synchronized = true;
      currentRaw = '';
      render();
    }
  };

  const startDraft = (seed: string, lease?: FormulaEditLease): boolean => {
    if (detached || activeDraft || !anchor) return activeDraft !== null;
    const handle = deps.beginDraft(
      { ...anchor },
      seed,
      {
        onFinish: (outcome, restored) => {
          if (binding) finishDraft(binding, outcome);
          if (!restored) return;
          // A restored edit owns the interaction again, so the palette session ends.
          restoredFocusTarget = restored;
          if (!closing && mode !== 'closed') close();
        },
      },
      lease ? { lease } : undefined,
    );
    if (!handle) return false;
    let binding!: DraftBinding;
    binding = {
      handle,
      unsubscribe: () => {},
      leased: lease !== undefined,
    };
    activeDraft = binding;
    currentRaw = handle.value();
    binding.unsubscribe = handle.subscribe((raw) => handleRawUpdate(binding, raw));
    notifyMirror(currentRaw);
    return true;
  };

  const ensurePostCommitDraft = (): boolean => {
    if (activeDraft) return true;
    if ((mode !== 'arguments-committed' && mode !== 'picker') || !postCommit) return false;
    if (!startDraft(postCommit.raw)) {
      args = [...postCommit.args];
      currentRaw = postCommit.raw;
      const parsed = selectedName ? parseOuterCall(currentRaw, selectedName) : null;
      explicitArgumentCount = parsed?.args.length ?? null;
      preserveExplicitArgumentCount = parsed !== null;
      synchronized = parsed !== null;
      guardMessage = labels.draftConflict;
      notifyMirror(null);
      render();
      return false;
    }
    mode = 'arguments-editing';
    const parsed = parseOuterCall(currentRaw, postCommit.name);
    synchronized = parsed !== null;
    explicitArgumentCount = parsed?.args.length ?? null;
    preserveExplicitArgumentCount = parsed !== null;
    guardMessage = '';
    updateDataset();
    if (argumentsRefs) {
      projectDisabledState(
        argumentsRefs.done,
        selectedUnavailable(),
        selectedUnavailable() ? labels.unavailable : null,
      );
    }
    return true;
  };

  const onArgumentInput = (index: number, value: string): void => {
    focusedArgument = index;
    if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
    if (mode !== 'arguments-editing' || !activeDraft || !synchronized) return;
    args[index] = value;
    if (preserveExplicitArgumentCount) {
      explicitArgumentCount = Math.max(explicitArgumentCount ?? 0, index + 1);
    }
    setDraftRaw(assembledCurrentFormula(selectedName ?? ''));
    renderPreview();
  };

  const argumentCount = (entry: FunctionCatalogEntry): number =>
    Math.max(entry.minArity, args.length, entry.argumentLabels.length);

  const renderArguments = (): void => {
    clearRoot();
    updateDataset();
    const header = document.createElement('header');
    header.className = 'fc-mac-formula-palette__header';
    const title = document.createElement('h2');
    title.textContent = labels.title;
    header.appendChild(title);
    header.appendChild(makeIconButton(labels.close, 'close', 'close', () => api.close()));
    root.appendChild(header);

    const content = document.createElement('div');
    content.className = 'fc-mac-formula-palette__content';
    const back = makeButton(labels.showAll, 'show-all', returnToPicker);
    back.className = 'fc-mac-formula-palette__back';
    content.appendChild(back);

    const name = document.createElement('h3');
    name.className = 'fc-mac-formula-palette__args-name';
    name.textContent = selectedEntry?.displayName ?? selectedName ?? '';
    content.appendChild(name);
    const fields = document.createElement('div');
    fields.className = 'fc-mac-formula-palette__fields';
    const count = selectedEntry ? argumentCount(selectedEntry) : args.length;
    for (let index = 0; index < count; index += 1) {
      const row = document.createElement('label');
      row.className = 'fc-mac-formula-palette__argument';
      const help = selectedEntry
        ? argumentHelp(selectedEntry, index)
        : { label: `${labels.argument} ${index + 1}` };
      const label = document.createElement('span');
      label.textContent = help.label ?? `${labels.argument} ${index + 1}`;
      row.appendChild(label);
      const input = document.createElement('input');
      input.type = 'text';
      input.dataset.argumentIndex = String(index);
      input.value = args[index] ?? '';
      input.addEventListener('focus', () => {
        focusedArgument = index;
      });
      input.addEventListener('input', () => onArgumentInput(index, input.value));
      row.appendChild(input);
      const range = makeIconButton(labels.rangePicker, 'range-picker', 'range', () => {
        focusedArgument = index;
        input.focus();
      });
      range.className = 'fc-range-picker__btn fc-mac-formula-palette__range';
      row.appendChild(range);
      if (help.description) {
        const hint = document.createElement('small');
        hint.textContent = help.description;
        row.appendChild(hint);
      }
      fields.appendChild(row);
    }
    content.appendChild(fields);
    if (selectedEntry?.maxArity === null) {
      const add = makeButton(labels.argument, 'add-argument', () => {
        if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
        if (!synchronized || !activeDraft || !selectedName) return;
        args.push('');
        preserveExplicitArgumentCount = true;
        explicitArgumentCount = Math.max(explicitArgumentCount ?? 0, args.length);
        const raw = assembledCurrentFormula(selectedName);
        setDraftRaw(raw, true);
        renderArguments();
      });
      add.className = 'fc-mac-formula-palette__add-argument';
      content.appendChild(add);
    }

    const resultRow = document.createElement('div');
    resultRow.className = 'fc-mac-formula-palette__result-row';
    const resultLabel = document.createElement('h4');
    resultLabel.textContent = labels.result;
    resultRow.appendChild(resultLabel);
    const preview = document.createElement('output');
    preview.className = 'fc-mac-formula-palette__preview';
    preview.dataset.role = 'preview-value';
    resultRow.appendChild(preview);
    const done = makeButton(labels.done, 'done', commitSelected);
    done.className = 'fc-mac-formula-palette__done';
    projectDisabledState(
      done,
      selectedUnavailable() || mode === 'arguments-committed',
      selectedUnavailable() || mode === 'arguments-committed' ? labels.unavailable : null,
    );
    resultRow.appendChild(done);
    content.appendChild(resultRow);
    if (guardMessage) {
      const guard = document.createElement('p');
      guard.className = 'fc-mac-formula-palette__guard';
      guard.dataset.role = 'draft-conflict';
      guard.textContent = guardMessage;
      content.appendChild(guard);
    }

    const help = document.createElement('section');
    help.className = 'fc-mac-formula-palette__help';
    const helpHeading = document.createElement('h4');
    helpHeading.textContent = labels.help;
    help.appendChild(helpHeading);
    const helpText = document.createElement('p');
    helpText.textContent = `${labels.description}: ${selectedDescription()}`;
    help.appendChild(helpText);
    const syntaxText = document.createElement('p');
    syntaxText.textContent = `${labels.syntax}: ${selectedSyntax()}`;
    help.appendChild(syntaxText);
    const firstArgumentHelp = selectedEntry ? argumentHelp(selectedEntry, 0) : null;
    if (firstArgumentHelp?.url) {
      const link = document.createElement('a');
      link.href = firstArgumentHelp.url;
      link.target = '_blank';
      link.rel = 'noreferrer';
      link.textContent = labels.help;
      help.appendChild(link);
    }
    content.appendChild(help);
    root.appendChild(content);
    argumentsRefs = { fields, preview, done };
    updateDataset();
    updateFieldValues();
    renderPreview();
  };

  const renderPicker = (): void => {
    clearRoot();
    updateDataset();
    const header = document.createElement('header');
    header.className = 'fc-mac-formula-palette__header';
    const title = document.createElement('h2');
    title.textContent = labels.title;
    header.appendChild(title);
    header.appendChild(makeIconButton(labels.close, 'close', 'close', () => api.close()));
    root.appendChild(header);

    const content = document.createElement('div');
    content.className = 'fc-mac-formula-palette__content';
    const search = document.createElement('input');
    search.type = 'search';
    search.className = 'fc-mac-formula-palette__search';
    search.placeholder = labels.searchPlaceholder;
    search.setAttribute('aria-label', labels.searchPlaceholder);
    search.value = searchQuery;
    search.addEventListener('input', () => {
      searchQuery = search.value;
      updatePickerList();
    });
    content.appendChild(search);
    const sections = document.createElement('div');
    sections.className = 'fc-mac-formula-palette__sections';
    content.appendChild(sections);
    const summary = document.createElement('div');
    summary.className = 'fc-mac-formula-palette__summary';
    content.appendChild(summary);
    const insert = makeButton(labels.insertFunction, 'insert-function', () =>
      insertSelected(pickerSelectionName ?? selectedName, pickerSelectionEntry ?? selectedEntry),
    );
    insert.className = 'fc-mac-formula-palette__insert';
    const pickerEntry = pickerSelectionEntry ?? selectedEntry;
    projectDisabledState(
      insert,
      (pickerSelectionName ?? selectedName) === null ||
        pickerEntry === null ||
        isFunctionUnavailableForInsertion(pickerEntry.availability),
      (pickerSelectionName ?? selectedName) === null ||
        pickerEntry === null ||
        isFunctionUnavailableForInsertion(pickerEntry.availability)
        ? labels.unavailable
        : null,
    );
    content.appendChild(insert);
    if (guardMessage) {
      const guard = document.createElement('p');
      guard.className = 'fc-mac-formula-palette__guard';
      guard.textContent = guardMessage;
      content.appendChild(guard);
    }
    root.appendChild(content);
    pickerRefs = { search, sections, insert };
    argumentsRefs = null;
    updatePickerSummary(summary);
    updatePickerList();
    updateDataset();
  };

  const render = (): void => {
    updateDataset();
    if (mode === 'closed') {
      clearRoot();
      pickerRefs = null;
      argumentsRefs = null;
      return;
    }
    if (mode === 'picker') renderPicker();
    else renderArguments();
  };

  const returnToPicker = (): void => {
    // A leased draft stays open so Close can still restore the suspended edit.
    const keptDraft = activeDraft?.leased ? activeDraft : null;
    if (activeDraft && !keptDraft) {
      closing = true;
      activeDraft.handle.cancel();
      closing = false;
    }
    selectedName = null;
    selectedEntry = null;
    args = [];
    explicitArgumentCount = null;
    preserveExplicitArgumentCount = false;
    pickerCategory = 'all';
    pickerSelectionName = null;
    pickerSelectionEntry = null;
    synchronized = true;
    currentRaw = '';
    guardMessage = '';
    mode = 'picker';
    if (keptDraft) setDraftRaw('=');
    else startDraft('=');
    render();
    pickerRefs?.search.focus();
  };

  const insertSelected = (
    requestedName: string | null = selectedName,
    requestedEntry: FunctionCatalogEntry | null = selectedEntry,
  ): void => {
    if (
      !requestedName ||
      !requestedEntry ||
      isFunctionUnavailableForInsertion(requestedEntry.availability)
    ) {
      guardMessage = labels.unavailable;
      render();
      return;
    }

    if (mode === 'picker' && activeDraft) {
      const raw = activeDraft.handle.value();
      const initialDraft = selectedName === null && raw.trim() === '=';
      const completeDraft =
        selectedName !== null && synchronized && parseOuterCall(raw, selectedName) !== null;
      if (!initialDraft && !completeDraft) {
        guardMessage = labels.draftConflict;
        synchronized = false;
        render();
        return;
      }
    } else if (mode === 'picker' && postCommit && !activeDraft) {
      if (!ensurePostCommitDraft() || !synchronized) {
        guardMessage = labels.draftConflict;
        render();
        return;
      }
    }

    selectedName = requestedName;
    selectedEntry = requestedEntry;
    pickerSelectionName = null;
    pickerSelectionEntry = null;
    if (!liveEntryGuard()) return;
    if (!activeDraft && !startDraft('=')) return;
    const initial = deps.getInitialArguments?.(selectedName) ?? [];
    const count = Math.max(
      selectedEntry.minArity,
      selectedEntry.argumentLabels.length,
      initial.length,
    );
    args = Array.from({ length: count }, (_, index) => initial[index] ?? '');
    explicitArgumentCount = null;
    preserveExplicitArgumentCount = false;
    synchronized = true;
    guardMessage = '';
    const formula = assembledFormula(selectedName, args);
    setDraftRaw(formula);
    mode = 'arguments-editing';
    postCommit = null;
    render();
    const first = argumentsRefs?.fields.querySelector<HTMLInputElement>(
      'input[data-argument-index="0"]',
    );
    first?.focus();
  };

  const focusCurrentSurface = (): void => {
    if (mode === 'picker') {
      pickerRefs?.search.focus();
      return;
    }
    const field =
      argumentsRefs?.fields.querySelector<HTMLInputElement>(
        `input[data-argument-index="${focusedArgument}"]`,
      ) ?? argumentsRefs?.fields.querySelector<HTMLInputElement>('input[data-argument-index]');
    field?.focus();
  };

  const rereadOpenCatalog = (): boolean => {
    strings = deps.getStrings();
    labels = paletteStrings(strings);
    const live = safeWorkbook();
    catalog = readCatalog();
    reconcilePickerSelection();
    if (live !== sessionWb) {
      showGuardedPicker(labels.unavailable);
      return false;
    }
    if (
      mode !== 'picker' &&
      selectedName &&
      (!catalog.knownNames.has(selectedName) ||
        !catalog.entries.has(selectedName) ||
        isFunctionUnavailableForInsertion(catalog.entries.get(selectedName)?.availability))
    ) {
      showGuardedPicker(labels.unavailable);
      return false;
    }
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
    return true;
  };

  const applyOpenRequest = (seedName?: string, options?: FxDialogOpenOptions): void => {
    if (!rereadOpenCatalog()) return;
    if (deps.store.getState().ui.editor.kind !== 'idle') {
      guardMessage = labels.draftConflict;
      render();
      return;
    }
    if (!seedName) {
      if (options?.category !== undefined) {
        pickerCategory = options.category;
        pickerSelectionName = null;
        pickerSelectionEntry = null;
        if (mode !== 'picker') mode = 'picker';
      }
      render();
      focusCurrentSurface();
      return;
    }

    const requestedName = canonicalName(seedName);
    const requestedEntry = catalog.entries.get(requestedName) ?? null;
    if (!requestedEntry || isFunctionUnavailableForInsertion(requestedEntry.availability)) {
      guardMessage = labels.unavailable;
      render();
      focusCurrentSurface();
      return;
    }

    if (mode === 'arguments-editing') {
      const raw = activeDraft?.handle.value() ?? currentRaw;
      if (
        !activeDraft ||
        !selectedName ||
        !synchronized ||
        parseOuterCall(raw, selectedName) === null
      ) {
        guardMessage = labels.unavailable;
        render();
        focusCurrentSurface();
        return;
      }
    } else if (mode === 'arguments-committed') {
      if (!ensurePostCommitDraft() || !synchronized) {
        guardMessage = labels.unavailable;
        render();
        focusCurrentSurface();
        return;
      }
    } else if (mode !== 'picker') {
      return;
    }

    guardMessage = '';
    insertSelected(requestedEntry.canonicalName, requestedEntry);
  };

  const commitSelected = (): void => {
    if (mode !== 'arguments-editing' || !activeDraft || !selectedName || !selectedEntry) return;
    if (selectedUnavailable()) {
      guardMessage = labels.unavailable;
      render();
      return;
    }
    const entry = liveEntryGuard();
    if (!entry || !activeDraft) return;
    const binding = activeDraft;
    committing = true;
    let committed = false;
    try {
      committed = binding.handle.commit();
    } finally {
      committing = false;
    }
    if (committed && activeDraft === binding) finishDraft(binding, 'committed');
  };

  const close = (): void => {
    if (detached || mode === 'closed') return;
    closing = true;
    if (activeDraft) activeDraft.handle.cancel();
    closing = false;
    mode = 'closed';
    selectedName = null;
    selectedEntry = null;
    args = [];
    explicitArgumentCount = null;
    preserveExplicitArgumentCount = false;
    pickerCategory = 'all';
    pickerSelectionName = null;
    pickerSelectionEntry = null;
    currentRaw = '';
    postCommit = null;
    synchronized = true;
    notifyMirror(null);
    render();
    const target = connectedElement(restoredFocusTarget) ?? connectedElement(opener) ?? deps.host;
    target.focus();
    restoredFocusTarget = null;
    opener = null;
    sessionWb = null;
    anchor = null;
  };

  const open = (seedName?: string, options?: FxDialogOpenOptions): void => {
    if (detached) return;
    if (mode !== 'closed') {
      applyOpenRequest(seedName, options);
      return;
    }
    opener = connectedElement(document.activeElement);
    anchor = { ...deps.getAnchor() };
    sessionWb = safeWorkbook();
    selectedName = null;
    selectedEntry = null;
    args = [];
    explicitArgumentCount = null;
    preserveExplicitArgumentCount = false;
    pickerCategory = 'all';
    pickerSelectionName = null;
    pickerSelectionEntry = null;
    synchronized = true;
    currentRaw = '';
    searchQuery = '';
    guardMessage = '';
    postCommit = null;
    restoredFocusTarget = null;
    mode = 'picker';
    const lease = deps.suspendActiveEdit?.(leaseContext) ?? undefined;
    if (!startDraft('=', lease) && lease) {
      // The draft refused the lease; hand the edit straight back to its owner.
      lease.userCancel()?.focus();
    }
    recentUnsubscribe?.();
    recentUnsubscribe = subscribeRecentFunctions(deps.store, () => {
      if (mode === 'picker') updatePickerList();
    });
    applyOpenRequest(seedName, options);
  };

  const refresh = (): void => {
    if (detached) return;
    strings = deps.getStrings();
    labels = paletteStrings(strings);
    catalog = readCatalog();
    reconcilePickerSelection();
    if (
      mode !== 'picker' &&
      selectedName &&
      (!catalog.knownNames.has(selectedName) ||
        !catalog.entries.has(selectedName) ||
        isFunctionUnavailableForInsertion(catalog.entries.get(selectedName)?.availability))
    ) {
      showGuardedPicker(labels.unavailable);
      return;
    }
    selectedEntry = selectedName ? (catalog.entries.get(selectedName) ?? null) : null;
    render();
  };

  const setStrings = (next: Strings): void => {
    strings = next;
    labels = paletteStrings(strings);
    render();
  };

  const rangeInsertTarget = (): RangeInsertTarget | null => {
    if ((mode !== 'arguments-editing' && mode !== 'arguments-committed') || !selectedName)
      return null;
    return {
      isFormulaEdit: () => Boolean((activeDraft || postCommit) && synchronized),
      insertRefAtCaret: (ref: string) => {
        if (mode === 'arguments-committed' && !ensurePostCommitDraft()) return;
        if (!activeDraft || !synchronized) return;
        args[focusedArgument] = ref;
        if (preserveExplicitArgumentCount) {
          explicitArgumentCount = Math.max(explicitArgumentCount ?? 0, focusedArgument + 1);
        }
        updateFieldValues();
        setDraftRaw(assembledCurrentFormula(selectedName ?? ''));
        renderPreview();
      },
    };
  };

  const detach = (): void => {
    if (detached) return;
    detached = true;
    closing = true;
    if (activeDraft) activeDraft.handle.cancel();
    closing = false;
    recentUnsubscribe?.();
    recentUnsubscribe = null;
    notifyMirror(null);
    deps.dock.hidden = true;
    root.remove();
    mode = 'closed';
    opener = null;
    restoredFocusTarget = null;
    anchor = null;
    sessionWb = null;
  };

  const api: MacFormulaPaletteHandle = {
    open,
    close,
    refresh,
    detach,
    isOpen: () => mode !== 'closed',
    rangeInsertTarget,
    setStrings,
  };
  return api;
}
