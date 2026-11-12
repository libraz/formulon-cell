import { canExecuteBuiltInRequest } from '../commands/built-in-command-policy.js';
import type { History } from '../commands/history.js';
import { namedRangeWrite } from '../commands/named-range-actions.js';
import {
  type DefinedNameEntry,
  definedNameDisplayFormula,
  definedNameExists,
  definedNameValueText,
  deleteDefinedName,
  isSameDefinedName,
  isValidDefinedName,
  sanitizeName,
  upsertDefinedName,
} from '../commands/named-ranges.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { handleMenuKeydown, projectDisabledState } from '../toolbar/menu-a11y.js';
import { formatSheetAbsoluteRange } from '../wrappers/toolbar-a1.js';
import {
  type DialogEntry,
  INSTANCE_API_DEFINE_NAME_ENTRY,
  INSTANCE_API_NAME_MANAGER_ENTRY,
} from './dialog-entry.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  appendDialogIconButton,
  clearDialogError,
  createDialogButton,
  createDialogShell,
  deferInitialFocus,
  focusAndSelectInput,
  showDialogError,
} from './dialog-shell.js';
import { onOutsidePress } from './outside-press.js';
import { overlayPortalFor } from './overlay-portal.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface NamedRangeDialogDeps {
  host: HTMLElement;
  /** Store whose interaction controller authorizes every defined-name change. */
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  history?: History | null;
  strings?: Strings;
  getSelectedRangeFormula?: () => string;
  /** Current selection; seeds the New Name form (Refers To and suggested name). */
  getSelection?: () => { active: Addr; range: Range };
  subscribeToRangeChanges?: (listener: () => void) => () => void;
  /** Called after a defined-name mutation writes through to the engine so the
   *  host can re-project recalculated cells into the store. */
  onAfterMutate?: () => void;
}

export interface NamedRangeDialogHandle {
  /** `entry` names the entry point the changes are authorized under; the instance API when absent. */
  open(entry?: DialogEntry): void;
  openNew(entry?: DialogEntry): void;
  close(): void;
  detach(): void;
  /** Swap to a fresh workbook handle (used when the host re-binds the engine). */
  bindWorkbook(next: WorkbookHandle): void;
}

type NameFilter = 'all' | 'errors' | 'noErrors' | 'workbook';
type NameSortKey = 'name' | 'value' | 'formula' | 'scope' | 'comment';

function createNameManagerButton(
  label: string,
  className: string,
  opts: { role?: string } = {},
): HTMLButtonElement {
  const button = createDialogButton({ label, baseClass: className });
  if (opts.role) button.setAttribute('role', opts.role);
  return button;
}

const hasFormulaError = (formula: string): boolean =>
  /#(?:REF|NAME|VALUE|DIV\/0|N\/A|NULL|NUM)!?/i.test(formula);

/** Map a defined-name mutation failure to its inline message. */
function errorMessageFor(
  reason: 'empty-name' | 'invalid-name' | 'empty-formula' | 'unsupported' | 'engine-failed',
  t: Strings['namedRangeDialog'],
  refusal: Strings['commandRefusal'],
): string {
  if (reason === 'invalid-name') return t.errorInvalidName;
  if (reason === 'empty-name') return t.errorEmptyName;
  if (reason === 'empty-formula') return t.errorEmptyFormula;
  return refusal.engineFailed;
}

/**
 * Name Manager: list of defined names plus Excel-like child dialogs for
 * New/Edit and a quick "Refers to" edit box. When the engine supports
 * `setDefinedName` (capability `definedNameMutate`), edits write through.
 * On engines that don't (the JS stub or older bundles), controls are read-only.
 */
export function attachNamedRangeDialog(deps: NamedRangeDialogDeps): NamedRangeDialogHandle {
  const { host } = deps;
  let wb = deps.wb;
  const history = deps.history ?? null;
  let dialogEntry: DialogEntry = INSTANCE_API_NAME_MANAGER_ENTRY;
  /** Whether the engine can store names and the opening entry may change them. */
  const mayMutate = (): boolean =>
    wb.capabilities.definedNameMutate &&
    canExecuteBuiltInRequest(deps.store, dialogEntry.commandId, dialogEntry.origin).allowed;
  /** Run a defined-name change as one authorized, undoable step under the opening entry. */
  const writeNames = <T>(write: () => T) =>
    namedRangeWrite(
      {
        store: deps.store,
        workbook: wb,
        history,
        origin: dialogEntry.origin,
        commandId: dialogEntry.commandId,
      },
      write,
    );
  const strings = deps.strings ?? defaultStrings;
  const shared = strings.namedRangeDialog;
  // The Mac Name Manager words a few labels differently and orders the
  // reference row and editor buttons its own way; resolved on every open.
  const isMac = (): boolean => isMacHost(host);
  let mac = isMac();
  let t: typeof shared = mac ? { ...shared, ...strings.mac.namedRangeDialog } : shared;
  const rangePickerOptions =
    deps.getSelectedRangeFormula && deps.subscribeToRangeChanges
      ? {
          label: strings.pivotTableDialog.rangePickerSelect,
          // A name always refers to a sheet-qualified absolute reference.
          getValue: (): string => {
            const selection = deps.getSelection?.();
            if (!selection) return deps.getSelectedRangeFormula?.() ?? '';
            const { range } = selection;
            return `=${formatSheetAbsoluteRange(wb.sheetName(range.sheet), range)}`;
          },
          subscribeToRangeChanges: deps.subscribeToRangeChanges,
        }
      : null;

  const shell = createDialogShell({
    host,
    className: 'fc-namedlg',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-namedlg__panel');
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  const actionBar = document.createElement('div');
  actionBar.className = 'fc-namedlg__actions';
  const newBtn = appendDialogButton(actionBar, { label: t.newButton });
  const editBtn = appendDialogButton(actionBar, { label: t.editButton });
  const deleteBtn = appendDialogButton(actionBar, { label: t.deleteButton });
  const filterBtn = appendDialogButton(actionBar, { label: t.filterButton });
  filterBtn.setAttribute('aria-haspopup', 'menu');
  body.appendChild(actionBar);

  const list = document.createElement('div');
  list.className = 'fc-namedlg__list';
  list.setAttribute('role', 'listbox');
  list.setAttribute('aria-label', t.title);
  const listHead = document.createElement('div');
  listHead.className = 'fc-namedlg__row fc-namedlg__head';
  const headerSpecs: { key: NameSortKey; label: string }[] = [
    { key: 'name', label: t.nameHeader },
    { key: 'value', label: t.valueHeader },
    { key: 'formula', label: t.formulaHeader },
    { key: 'scope', label: t.scopeHeader },
    { key: 'comment', label: t.commentHeader },
  ];
  const headerButtons = new Map<NameSortKey, HTMLButtonElement>();
  for (const { key, label } of headerSpecs) {
    const cell = document.createElement('span');
    const button = createNameManagerButton(label, 'fc-namedlg__sort');
    headerButtons.set(key, button);
    button.dataset.sortKey = key;
    button.addEventListener('click', () => {
      if (sortKey === key) sortDir = sortDir === 'asc' ? 'desc' : 'asc';
      else {
        sortKey = key;
        sortDir = 'asc';
      }
      selectedNameIndex = 0;
      renderList();
    });
    cell.appendChild(button);
    listHead.appendChild(cell);
  }
  list.appendChild(listHead);
  body.appendChild(list);

  const quickRow = document.createElement('div');
  quickRow.className = 'fc-namedlg__refers';
  const quickLabel = document.createElement('label');
  quickLabel.className = 'fc-namedlg__refers-label';
  const quickLabelText = document.createElement('span');
  quickLabelText.className = 'fc-namedlg__refers-caption';
  quickLabelText.textContent = t.quickRefersLabel;
  quickLabel.appendChild(quickLabelText);
  const quickInput = document.createElement('input');
  quickInput.type = 'text';
  quickInput.className = 'fc-namedlg__refers-input';
  quickInput.setAttribute('aria-label', t.formulaHeader);
  quickInput.autocomplete = 'off';
  quickInput.spellcheck = false;
  quickLabel.appendChild(quickInput);
  if (rangePickerOptions) {
    attachRangePickerButton(quickInput, {
      ...rangePickerOptions,
      kind: 'named-range-quick-refers-to',
    });
  }
  quickRow.appendChild(quickLabel);
  const quickCommitBtn = appendDialogIconButton(quickRow, {
    label: '',
    ariaLabel: t.commitButton,
    baseClass: 'fc-namedlg__refers-icon',
  });
  quickCommitBtn.classList.add('fc-namedlg__refers-icon--commit');
  const quickCancelBtn = appendDialogIconButton(quickRow, {
    label: '',
    ariaLabel: t.cancelButton,
    baseClass: 'fc-namedlg__refers-icon',
  });
  quickCancelBtn.classList.add('fc-namedlg__refers-icon--cancel');
  body.appendChild(quickRow);

  const editorShell = createDialogShell({
    host,
    className: 'fc-namedlg-editor',
    ariaLabel: t.newNameTitle,
    onDismiss: () => closeEditor(),
  });
  editorShell.overlay.classList.add('fc-fmtdlg');
  editorShell.panel.classList.add('fc-fmtdlg__panel', 'fc-namedlg-editor__panel');
  const editorTitle = document.createElement('div');
  editorTitle.className = 'fc-fmtdlg__header';
  editorShell.panel.appendChild(editorTitle);
  const editorBody = document.createElement('div');
  editorBody.className = 'fc-namedlg-editor__body';
  editorShell.panel.appendChild(editorBody);

  const formRow = document.createElement('form');
  formRow.className = 'fc-namedlg-editor__form';
  const makeEditorRow = (labelText: string, control: HTMLElement): HTMLLabelElement => {
    const label = document.createElement('label');
    label.className = 'fc-namedlg-editor__row';
    const text = document.createElement('span');
    text.textContent = labelText;
    label.append(text, control);
    return label;
  };
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.className = 'fc-namedlg-editor__input';
  nameInput.placeholder = t.namePlaceholder;
  nameInput.setAttribute('aria-label', t.nameHeader);
  nameInput.autocomplete = 'off';
  nameInput.spellcheck = false;
  const scopeSelect = createDialogSelect([], '-1', {
    className: 'fc-namedlg-editor__input',
    ariaLabel: t.scopeHeader,
  });
  const commentInput = document.createElement('textarea');
  commentInput.className = 'fc-namedlg-editor__input';
  commentInput.setAttribute('aria-label', t.commentHeader);
  commentInput.rows = 2;
  const formulaInput = document.createElement('input');
  formulaInput.type = 'text';
  formulaInput.className = 'fc-namedlg-editor__input';
  formulaInput.placeholder = t.formulaPlaceholder;
  formulaInput.setAttribute('aria-label', t.formulaHeader);
  formulaInput.autocomplete = 'off';
  formulaInput.spellcheck = false;
  const formulaRow = makeEditorRow(t.formulaLabel, formulaInput);
  formRow.append(
    makeEditorRow(t.nameLabel, nameInput),
    makeEditorRow(t.scopeLabel, scopeSelect),
    makeEditorRow(t.commentLabel, commentInput),
    formulaRow,
  );
  if (rangePickerOptions) {
    attachRangePickerButton(formulaInput, {
      ...rangePickerOptions,
      kind: 'named-range-editor-refers-to',
    });
  }
  // While the editor's field picks from the grid, the Name Manager beneath must not catch the clicks.
  const syncManagerPassthrough = (): void => {
    const cls = editorShell.overlay.classList;
    overlay.style.pointerEvents =
      cls.contains('fc-fmtdlg--range-armed') || cls.contains('fc-fmtdlg--range-picking')
        ? 'none'
        : '';
  };
  const passthroughObserver = new MutationObserver(syncManagerPassthrough);
  passthroughObserver.observe(editorShell.overlay, {
    attributes: true,
    attributeFilter: ['class'],
  });
  const editorButtons = document.createElement('div');
  editorButtons.className = 'fc-namedlg-editor__buttons';
  const editorOkBtn = appendDialogButton(editorButtons, { label: t.ok, variant: 'primary' });
  editorOkBtn.type = 'submit';
  const editorCancelBtn = appendDialogButton(editorButtons, { label: t.cancel });
  formRow.appendChild(editorButtons);

  const applyPlatform = (): void => {
    mac = isMac();
    t = mac ? { ...shared, ...strings.mac.namedRangeDialog } : shared;
    header.textContent = t.title;
    shell.setAriaLabel(t.title);
    list.setAttribute('aria-label', t.title);
    newBtn.textContent = t.newButton;
    const formulaHeaderBtn = headerButtons.get('formula');
    if (formulaHeaderBtn) formulaHeaderBtn.textContent = t.formulaHeader;
    quickLabelText.textContent = t.quickRefersLabel;
    quickInput.setAttribute('aria-label', t.formulaHeader);
    formulaInput.setAttribute('aria-label', t.formulaHeader);
    (formulaRow.firstElementChild as HTMLElement).textContent = t.formulaLabel;
    if (mac) {
      quickRow.append(quickCancelBtn, quickCommitBtn, quickLabel);
      editorButtons.append(editorCancelBtn, editorOkBtn);
    } else {
      quickRow.append(quickLabel, quickCommitBtn, quickCancelBtn);
      editorButtons.append(editorOkBtn, editorCancelBtn);
    }
  };
  const errorRow = document.createElement('div');
  errorRow.className = 'fc-namedlg__error';
  errorRow.setAttribute('role', 'alert');
  errorRow.hidden = true;
  editorBody.append(formRow, errorRow);

  const deleteShell = createDialogShell({
    host,
    className: 'fc-namedlg-confirm',
    ariaLabel: t.title,
    onDismiss: () => closeDeleteConfirm(),
  });
  deleteShell.overlay.classList.add('fc-fmtdlg');
  deleteShell.panel.classList.add('fc-fmtdlg__panel', 'fc-namedlg-confirm__panel');
  const deleteTitle = document.createElement('div');
  deleteTitle.className = 'fc-fmtdlg__header';
  deleteTitle.textContent = t.title;
  const deleteBody = document.createElement('div');
  deleteBody.className = 'fc-namedlg-confirm__body';
  const deleteMessage = document.createElement('p');
  deleteMessage.className = 'fc-namedlg-confirm__message';
  const deleteButtons = document.createElement('div');
  deleteButtons.className = 'fc-namedlg-confirm__buttons';
  const deleteOkBtn = appendDialogButton(deleteButtons, { label: t.ok, variant: 'primary' });
  const deleteCancelBtn = appendDialogButton(deleteButtons, { label: t.cancel });
  deleteBody.append(deleteMessage, deleteButtons);
  deleteShell.panel.append(deleteTitle, deleteBody);

  const mainErrorRow = document.createElement('div');
  mainErrorRow.className = 'fc-namedlg__error';
  mainErrorRow.setAttribute('role', 'alert');
  mainErrorRow.hidden = true;
  body.appendChild(mainErrorRow);

  const note = document.createElement('div');
  note.className = 'fc-namedlg__note';
  note.id = 'fc-namedlg-readonly-note';
  note.textContent = t.note;

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);
  const closeBtn = appendDialogButton(footer, { label: t.close });

  const showError = (msg: string): void => {
    showDialogError(errorRow, msg);
  };
  const clearError = (): void => {
    clearDialogError(errorRow);
  };
  const showMainError = (msg: string): void => {
    showDialogError(mainErrorRow, msg);
  };
  const clearMainError = (): void => {
    clearDialogError(mainErrorRow);
  };
  let selectedNameIndex = 0;
  let currentRows: DefinedNameEntry[] = [];
  let activeFilter: NameFilter = 'all';
  let sortKey: NameSortKey = 'name';
  let sortDir: 'asc' | 'desc' = 'asc';
  let filterMenu: HTMLDivElement | null = null;
  let detachFilterOutside: (() => void) | null = null;
  let editorMode: 'new' | 'edit' = 'new';
  let pendingDeleteName: DefinedNameEntry | null = null;
  let editingEntry: DefinedNameEntry | null = null;

  const setReadOnlyState = (
    el: HTMLElement & { disabled?: boolean },
    disabledByReadOnly: boolean,
  ): void => {
    if (disabledByReadOnly) delete el.dataset.disabledReason;
    projectDisabledState(el, disabledByReadOnly, t.note, {
      ariaDescription: false,
      describedById: note.id,
    });
  };
  const setStateDisabledReason = (
    el: HTMLElement & { disabled?: boolean },
    disabled: boolean,
    reason: string | null,
    titlePrefix: string,
  ): void => {
    projectDisabledState(el, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix,
    });
  };

  const updateRowSelection = (): HTMLElement[] => {
    const rows = Array.from(list.querySelectorAll<HTMLElement>('.fc-namedlg__item'));
    if (rows.length === 0) return rows;
    selectedNameIndex = Math.min(selectedNameIndex, rows.length - 1);
    for (const [rowIdx, row] of rows.entries()) {
      const selected = rowIdx === selectedNameIndex;
      row.tabIndex = selected ? 0 : -1;
      row.setAttribute('aria-selected', selected ? 'true' : 'false');
      row.classList.toggle('fc-namedlg__item--selected', selected);
    }
    syncQuickRefers();
    return rows;
  };

  const focusNameRow = (idx: number): void => {
    const rows = Array.from(list.querySelectorAll<HTMLElement>('.fc-namedlg__item'));
    if (rows.length === 0) return;
    selectedNameIndex = (idx + rows.length) % rows.length;
    updateRowSelection();
    rows[selectedNameIndex]?.focus({ preventScroll: true });
  };

  const syncQuickRefers = (): void => {
    const entry = currentRows[selectedNameIndex];
    const canMutate = mayMutate();
    const enabled = canMutate && Boolean(entry);
    quickInput.value = definedNameDisplayFormula(entry?.formula ?? '');
    if (!canMutate) {
      setReadOnlyState(quickInput, true);
      setReadOnlyState(quickCommitBtn, true);
      setReadOnlyState(quickCancelBtn, true);
    } else {
      const reason = entry ? null : t.quickRefersRequiresSelection;
      setStateDisabledReason(quickInput, !enabled, reason, t.formulaHeader);
      setStateDisabledReason(quickCommitBtn, !enabled, reason, t.commitButton);
      setStateDisabledReason(quickCancelBtn, !entry, reason, t.cancelButton);
    }
  };

  const renderList = (): void => {
    list.replaceChildren();
    list.appendChild(listHead);
    for (const button of listHead.querySelectorAll<HTMLButtonElement>('.fc-namedlg__sort')) {
      const selected = button.dataset.sortKey === sortKey;
      button.setAttribute(
        'aria-sort',
        selected ? (sortDir === 'asc' ? 'ascending' : 'descending') : 'none',
      );
      button.classList.toggle('fc-namedlg__sort--active', selected);
      button.dataset.sortDir = selected ? sortDir : '';
    }
    const canMutate = mayMutate();
    const allRows = [...wb.definedNames()].filter((entry) => !entry.hidden);
    currentRows = allRows.filter((entry) => {
      if (activeFilter === 'errors') return hasFormulaError(entry.formula);
      if (activeFilter === 'noErrors') return !hasFormulaError(entry.formula);
      if (activeFilter === 'workbook') return entry.localSheetId < 0;
      return true;
    });
    currentRows.sort((a, b) => {
      const left = sortValue(a, sortKey);
      const right = sortValue(b, sortKey);
      const result = left.localeCompare(right, undefined, { numeric: true, sensitivity: 'base' });
      return sortDir === 'asc' ? result : -result;
    });
    const count = currentRows.length;
    for (const [rowIndex, entry] of currentRows.entries()) {
      const item = document.createElement('div');
      item.className = 'fc-namedlg__item fc-namedlg__row';
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', rowIndex === selectedNameIndex ? 'true' : 'false');
      item.tabIndex = rowIndex === selectedNameIndex ? 0 : -1;
      item.classList.toggle('fc-namedlg__item--selected', rowIndex === selectedNameIndex);
      const name = document.createElement('span');
      name.className = 'fc-namedlg__row-name';
      name.textContent = entry.name;
      const value = document.createElement('span');
      value.className = 'fc-namedlg__row-value';
      value.textContent = definedNameValueText(wb, entry);
      const formulaCell = document.createElement('span');
      formulaCell.className = 'fc-namedlg__row-formula';
      formulaCell.textContent = definedNameDisplayFormula(entry.formula);
      const scope = document.createElement('span');
      scope.className = 'fc-namedlg__row-scope';
      scope.textContent = scopeLabel(entry.localSheetId);
      const comment = document.createElement('span');
      comment.className = 'fc-namedlg__row-comment';
      comment.textContent = entry.comment ?? '';
      item.addEventListener('click', () => {
        selectedNameIndex = rowIndex;
        updateRowSelection();
      });
      item.addEventListener('dblclick', () => {
        if (canMutate) editBtn.click();
      });
      item.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          focusNameRow(rowIndex + 1);
        } else if (e.key === 'ArrowUp') {
          e.preventDefault();
          focusNameRow(rowIndex - 1);
        } else if (e.key === 'Home') {
          e.preventDefault();
          focusNameRow(0);
        } else if (e.key === 'End') {
          e.preventDefault();
          focusNameRow(count - 1);
        } else if ((e.key === 'Delete' || e.key === 'Backspace') && canMutate) {
          e.preventDefault();
          requestDeleteSelectedName(item);
        } else if ((e.key === 'Enter' || e.key === ' ') && canMutate) {
          e.preventDefault();
          editBtn.click();
        }
      });
      item.append(name, value, formulaCell, scope, comment);
      list.appendChild(item);
    }
    if (count === 0) {
      selectedNameIndex = 0;
      const empty = document.createElement('div');
      empty.className = 'fc-namedlg__empty';
      // The Mac list stays blank when no names exist.
      empty.textContent = mac ? '' : t.empty;
      list.appendChild(empty);
    }
    updateRowSelection();
    const hasSelection = count > 0;
    if (!canMutate) {
      setReadOnlyState(editBtn, true);
      setReadOnlyState(deleteBtn, true);
    } else {
      const selectionReason = hasSelection ? null : t.selectNameActionReason;
      setStateDisabledReason(editBtn, !hasSelection, selectionReason, t.editButton);
      setStateDisabledReason(deleteBtn, !hasSelection, selectionReason, t.deleteButton);
    }
    setStateDisabledReason(
      filterBtn,
      allRows.length === 0,
      allRows.length === 0 ? t.filterRequiresNames : null,
      t.filterButton,
    );
    syncQuickRefers();
    filterBtn.textContent =
      activeFilter === 'all' ? t.filterButton : `${t.filterButton}: ${filterLabel(activeFilter)}`;
  };

  const sortValue = (entry: DefinedNameEntry, key: NameSortKey): string => {
    switch (key) {
      case 'formula':
        return entry.formula;
      case 'scope':
        return scopeLabel(entry.localSheetId);
      case 'value':
        return definedNameValueText(wb, entry);
      case 'comment':
        return entry.comment ?? '';
      default:
        return entry.name;
    }
  };

  const scopeLabel = (localSheetId: number): string =>
    localSheetId < 0 ? t.workbookScope : wb.sheetName(localSheetId);

  const scopeOptions = (): { value: string; label: string }[] => [
    { value: '-1', label: t.workbookScope },
    ...Array.from({ length: wb.sheetCount }, (_, sheet) => ({
      value: String(sheet),
      label: wb.sheetName(sheet),
    })),
  ];

  const selectedScopeId = (): number => {
    const value = Number(scopeSelect.value);
    return Number.isInteger(value) && value >= -1 ? value : -1;
  };

  const filterLabel = (filter: NameFilter): string => {
    switch (filter) {
      case 'errors':
        return t.filterNamesWithErrors;
      case 'noErrors':
        return t.filterNamesWithoutErrors;
      case 'workbook':
        return t.filterNamesScopedToWorkbook;
      default:
        return t.filterAll;
    }
  };

  const closeFilterMenu = (): void => {
    filterMenu?.remove();
    filterMenu = null;
    detachFilterOutside?.();
    detachFilterOutside = null;
    document.removeEventListener('keydown', onFilterDocKey, true);
  };

  function onFilterDocKey(e: KeyboardEvent): void {
    if (!filterMenu) return;
    handleMenuKeydown(e, filterMenu, {
      close: closeFilterMenu,
      restoreFocusTo: filterBtn,
      items: '[role="menuitem"]',
    });
  }

  const openFilterMenu = (): void => {
    closeFilterMenu();
    if (filterBtn.disabled) return;
    const menu = document.createElement('div');
    menu.className = 'fc-namedlg__filter-menu';
    menu.setAttribute('role', 'menu');
    const filters: NameFilter[] = ['all', 'errors', 'noErrors', 'workbook'];
    for (const filter of filters) {
      const item = createNameManagerButton(filterLabel(filter), 'fc-namedlg__filter-item', {
        role: 'menuitem',
      });
      item.setAttribute('aria-checked', filter === activeFilter ? 'true' : 'false');
      item.addEventListener('click', () => {
        activeFilter = filter;
        selectedNameIndex = 0;
        closeFilterMenu();
        renderList();
        filterBtn.focus();
      });
      menu.appendChild(item);
    }
    overlayPortalFor(host).appendChild(menu);
    const r = filterBtn.getBoundingClientRect();
    menu.style.left = `${Math.max(4, r.left)}px`;
    menu.style.top = `${r.bottom + 2}px`;
    filterMenu = menu;
    detachFilterOutside = onOutsidePress(
      document,
      () => [filterMenu, filterBtn],
      () => closeFilterMenu(),
      { event: 'pointerdown' },
    );
    document.addEventListener('keydown', onFilterDocKey, true);
    menu.querySelector<HTMLButtonElement>('[role="menuitem"]')?.focus();
  };

  function closeEditor(): void {
    editorShell.overlay.dispatchEvent(new CustomEvent('fc-range-picker-stop-all'));
    editorShell.close();
    syncManagerPassthrough();
    clearError();
    if (!shell.isOpen()) host.focus();
  }

  const openEditor = (mode: 'new' | 'edit'): void => {
    const entry = mode === 'edit' ? currentRows[selectedNameIndex] : null;
    if (mode === 'edit' && !entry) return;
    overlay.dispatchEvent(new CustomEvent('fc-range-picker-stop-all'));
    editorMode = mode;
    const title = mode === 'new' ? t.newNameTitle : t.editNameTitle;
    editorTitle.textContent = title;
    editorShell.setAriaLabel(title);
    scopeSelect.replaceChildren();
    appendDialogSelectOptions(scopeSelect, scopeOptions());
    editingEntry = entry ?? null;
    projectDisabledState(scopeSelect, mode === 'edit', null);
    scopeSelect.value = String(entry?.localSheetId ?? -1);
    commentInput.value = entry?.comment ?? '';
    if (entry) {
      nameInput.value = entry.name;
      formulaInput.value = definedNameDisplayFormula(entry.formula);
    } else {
      const seed = newNameSeed();
      nameInput.value = seed.name;
      formulaInput.value = seed.formula;
    }
    clearError();
    if (shell.isOpen()) {
      (mode === 'new' ? newBtn : editBtn).focus({ preventScroll: true });
    }
    editorShell.open();
    deferInitialFocus(editorShell.overlay, () => {
      focusAndSelectInput(nameInput);
    });
  };

  const newNameSeed = (): { name: string; formula: string } => {
    const selection = deps.getSelection?.();
    if (!selection) return { name: '', formula: deps.getSelectedRangeFormula?.() ?? '' };
    const { active, range } = selection;
    const formula = `=${formatSheetAbsoluteRange(wb.sheetName(range.sheet), range)}`;
    const textAt = (row: number, col: number): string => {
      if (row < 0 || col < 0) return '';
      const v = wb.getValue({ sheet: active.sheet, row, col });
      return v.kind === 'text' ? v.value.trim() : '';
    };
    const suggest = (text: string): string => {
      if (!text) return '';
      if (isValidDefinedName(text)) return text;
      const cleaned = sanitizeName(text, '');
      return isValidDefinedName(cleaned) ? cleaned : '';
    };
    const name =
      suggest(textAt(active.row, active.col)) ||
      suggest(textAt(active.row, active.col - 1)) ||
      suggest(textAt(active.row - 1, active.col));
    return { name, formula };
  };

  const loadSelectedNameIntoForm = (): void => {
    const entry = currentRows[selectedNameIndex];
    if (!entry) return;
    openEditor('edit');
  };

  function closeDeleteConfirm(): void {
    deleteShell.close();
    pendingDeleteName = null;
  }

  const requestDeleteSelectedName = (opener: HTMLElement = deleteBtn): void => {
    const entry = currentRows[selectedNameIndex];
    if (!entry || deleteBtn.disabled) return;
    pendingDeleteName = entry;
    deleteMessage.textContent = t.confirmDeleteMessage.replace('{name}', entry.name);
    opener.focus({ preventScroll: true });
    deleteShell.open();
  };

  // ── Mac 名前の定義: names list with + / -, a name and a range field ──
  const defineShell = createDialogShell({
    host,
    className: 'fc-namedlg-define',
    ariaLabel: t.title,
    onDismiss: () => closeDefine(),
  });
  defineShell.overlay.classList.add('fc-fmtdlg');
  const defineFrame = appendDialogFrame(defineShell, {
    title: '',
    panelClasses: ['fc-fmtdlg__panel', 'fc-namedlg-define__panel'],
    bodyClass: 'fc-fmtdlg__body fc-namedlg-define__body',
    footerClass: 'fc-fmtdlg__footer fc-namedlg-define__footer',
  });
  const defineNamesLabel = document.createElement('div');
  defineNamesLabel.className = 'fc-namedlg-define__names-label';
  const defineNames = document.createElement('div');
  defineNames.className = 'fc-namedlg-define__names';
  defineNames.setAttribute('role', 'listbox');
  const defineListButtons = document.createElement('div');
  defineListButtons.className = 'fc-namedlg-define__list-buttons';
  const defineAddBtn = appendDialogIconButton(defineListButtons, {
    label: '+',
    ariaLabel: t.addButton,
    baseClass: 'fc-namedlg-define__list-btn',
  });
  defineAddBtn.dataset.defineAction = 'add';
  const defineRemoveBtn = appendDialogIconButton(defineListButtons, {
    label: '−',
    ariaLabel: t.deleteButton,
    baseClass: 'fc-namedlg-define__list-btn',
  });
  defineRemoveBtn.dataset.defineAction = 'remove';
  const defineLeft = document.createElement('div');
  defineLeft.className = 'fc-namedlg-define__left';
  defineLeft.append(defineNamesLabel, defineNames, defineListButtons);
  const defineField = (): {
    label: HTMLLabelElement;
    text: HTMLSpanElement;
    input: HTMLInputElement;
  } => {
    const label = document.createElement('label');
    label.className = 'fc-namedlg-define__field';
    const text = document.createElement('span');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'fc-namedlg-editor__input';
    input.autocomplete = 'off';
    input.spellcheck = false;
    label.append(text, input);
    return { label, text, input };
  };
  const defineName = defineField();
  const defineRange = defineField();
  const defineRight = document.createElement('div');
  defineRight.className = 'fc-namedlg-define__right';
  defineRight.append(defineName.label, defineRange.label);
  if (rangePickerOptions) {
    attachRangePickerButton(defineRange.input, {
      ...rangePickerOptions,
      kind: 'named-range-define-refers-to',
    });
  }
  const defineError = document.createElement('div');
  defineError.className = 'fc-namedlg__error';
  defineError.setAttribute('role', 'alert');
  defineError.hidden = true;
  defineFrame.body.append(defineLeft, defineRight, defineError);
  const defineActions = appendDialogActions(defineFrame.footer, {
    cancelLabel: t.close,
    okLabel: t.ok,
  });
  let defineSelected: DefinedNameEntry | null = null;

  const renderDefineNames = (): void => {
    const entries = [...wb.definedNames()].filter((entry) => !entry.hidden);
    defineNames.replaceChildren(
      ...entries.map((entry) => {
        const option = document.createElement('div');
        option.setAttribute('role', 'option');
        option.dataset.name = entry.name;
        option.textContent = entry.name;
        const selected =
          defineSelected !== null && isSameDefinedName(defineSelected.name, entry.name);
        option.setAttribute('aria-selected', String(selected));
        option.addEventListener('click', () => {
          defineSelected = entry;
          defineName.input.value = entry.name;
          defineRange.input.value = definedNameDisplayFormula(entry.formula);
          renderDefineNames();
        });
        return option;
      }),
    );
    projectDisabledState(defineRemoveBtn, defineSelected === null, null);
  };
  const showDefineError = (message: string): void => showDialogError(defineError, message);
  /** Add or update the name in the fields; false when it was refused. */
  const defineCommit = (): boolean => {
    const name = defineName.input.value.trim();
    if (!name) {
      showDefineError(t.errorEmptyName);
      focusAndSelectInput(defineName.input);
      return false;
    }
    const written = writeNames(() =>
      upsertDefinedName(wb, name, defineRange.input.value.trim(), -1),
    );
    if (!written.ok) return false;
    const result = written.value;
    if (!result.ok) {
      showDefineError(errorMessageFor(result.reason, t, strings.commandRefusal));
      return false;
    }
    clearDialogError(defineError);
    deps.onAfterMutate?.();
    defineSelected = null;
    renderDefineNames();
    return true;
  };
  const closeDefine = (): void => {
    defineShell.close();
    host.focus();
  };
  defineShell.on(defineAddBtn, 'click', () => {
    defineCommit();
  });
  defineShell.on(defineRemoveBtn, 'click', () => {
    const entry = defineSelected;
    if (!entry) return;
    const written = writeNames(() => deleteDefinedName(wb, entry.name, entry.localSheetId));
    if (!written.ok) return;
    if (!written.value.ok) {
      showDefineError(strings.commandRefusal.engineFailed);
      return;
    }
    clearDialogError(defineError);
    deps.onAfterMutate?.();
    defineSelected = null;
    defineName.input.value = '';
    renderDefineNames();
  });
  defineShell.on(defineActions.cancelBtn, 'click', closeDefine);
  defineShell.on(defineActions.okBtn, 'click', () => {
    if (defineName.input.value.trim() && !defineCommit()) return;
    closeDefine();
  });
  const openDefine = (): void => {
    const m = strings.mac.namedRangeDialog;
    defineFrame.header.textContent = m.defineTitle;
    defineShell.setAriaLabel(m.defineTitle);
    defineNamesLabel.textContent = m.defineNames;
    defineName.text.textContent = m.defineNamePrompt;
    defineRange.text.textContent = m.defineRangePrompt;
    defineSelected = null;
    defineName.input.value = '';
    // The range seeds as an absolute, sheet-qualified reference: =Sheet!$A$1:$D$7.
    defineRange.input.value = newNameSeed().formula;
    clearDialogError(defineError);
    renderDefineNames();
    defineShell.open();
    deferInitialFocus(defineShell.overlay, () => defineName.input.focus());
  };

  const confirmDeleteSelectedName = (): void => {
    if (!pendingDeleteName) return;
    const entry = pendingDeleteName;
    const written = writeNames(() => deleteDefinedName(wb, entry.name, entry.localSheetId));
    if (!written.ok || !written.value.ok) {
      showMainError(strings.commandRefusal.engineFailed);
      closeDeleteConfirm();
      return;
    }
    clearMainError();
    deps.onAfterMutate?.();
    closeDeleteConfirm();
    selectedNameIndex = Math.min(selectedNameIndex, Math.max(0, currentRows.length - 2));
    renderList();
    if (!shell.isOpen()) return;
    const rows = list.querySelectorAll<HTMLElement>('.fc-namedlg__item');
    if (rows.length > 0) focusNameRow(selectedNameIndex);
    else if (!newBtn.disabled) newBtn.focus({ preventScroll: true });
    else closeBtn.focus({ preventScroll: true });
  };

  const commitQuickRefers = (): void => {
    const entry = currentRows[selectedNameIndex];
    if (!entry) return;
    const written = writeNames(() =>
      upsertDefinedName(wb, entry.name, quickInput.value, entry.localSheetId),
    );
    if (!written.ok) return;
    const result = written.value;
    if (!result.ok) {
      showMainError(errorMessageFor(result.reason, t, strings.commandRefusal));
      return;
    }
    clearMainError();
    deps.onAfterMutate?.();
    renderList();
    focusNameRow(selectedNameIndex);
  };

  const cancelQuickRefers = (): void => {
    syncQuickRefers();
    clearMainError();
    quickInput.focus();
  };

  /** Add `name`; when renaming, drop the entry it replaces so no copy is left behind. */
  const replaceDefinedName = (
    previous: DefinedNameEntry | null,
    name: string,
    formula: string,
    scope: number,
    comment: string,
  ): ReturnType<typeof upsertDefinedName> => {
    const renamed = previous !== null && previous.name !== name.trim();
    // A case-only rename addresses the same engine entry, so the old one goes first.
    const caseOnly = renamed && isSameDefinedName(previous.name, name);
    if (renamed && caseOnly) deleteDefinedName(wb, previous.name, scope);
    const result = upsertDefinedName(wb, name, formula, scope, { comment });
    if (!result.ok && renamed && caseOnly) {
      upsertDefinedName(wb, previous.name, previous.formula, scope, {
        comment: previous.comment ?? '',
      });
    }
    if (renamed && !caseOnly) {
      if (result.ok && !deleteDefinedName(wb, previous.name, scope).ok) {
        deleteDefinedName(wb, result.entry.name, scope);
        return { ok: false, reason: 'engine-failed' };
      }
    }
    return result;
  };

  const onSubmit = (e: SubmitEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    const name = nameInput.value.trim();
    const formula = formulaInput.value.trim();
    if (!name) {
      showError(t.errorEmptyName);
      focusAndSelectInput(nameInput);
      return;
    }
    const scope = selectedScopeId();
    const previous = editorMode === 'edit' ? editingEntry : null;
    const isSameEntry = previous !== null && isSameDefinedName(previous.name, name);
    if (!isSameEntry && definedNameExists(wb, name, scope)) {
      showError(t.errorDuplicateName);
      focusAndSelectInput(nameInput);
      return;
    }
    const written = writeNames(() =>
      replaceDefinedName(previous, name, formula, scope, commentInput.value.trim()),
    );
    if (!written.ok) return;
    const result = written.value;
    if (!result.ok) {
      showError(errorMessageFor(result.reason, t, strings.commandRefusal));
      return;
    }
    clearError();
    clearMainError();
    deps.onAfterMutate?.();
    nameInput.value = '';
    formulaInput.value = '';
    commentInput.value = '';
    editorShell.close();
    renderList();
    if (!shell.isOpen()) host.focus();
    else if (editorMode === 'new') newBtn.focus();
    else focusNameRow(selectedNameIndex);
  };

  const refreshFormState = (): void => {
    const canMutate = mayMutate();
    if (!canMutate) {
      setReadOnlyState(filterBtn, currentRows.length === 0);
    }
    setReadOnlyState(newBtn, !canMutate);
    setReadOnlyState(editBtn, !canMutate);
    setReadOnlyState(deleteBtn, !canMutate);
    setReadOnlyState(quickInput, !canMutate);
    setReadOnlyState(quickCommitBtn, !canMutate);
    // The engine must be able to store the note, or the field would silently drop it.
    setReadOnlyState(commentInput, !wb.capabilities.definedNameAttrs);
    if (canMutate) {
      setStateDisabledReason(
        filterBtn,
        currentRows.length === 0,
        currentRows.length === 0 ? t.filterRequiresNames : null,
        t.filterButton,
      );
    }
    if (canMutate) {
      if (note.parentElement) note.remove();
    } else {
      if (!note.parentElement) body.appendChild(note);
    }
    clearError();
    clearMainError();
  };

  const onClose = (): void => api.close();
  const onNew = (): void => {
    openEditor('new');
  };
  const onEdit = (): void => loadSelectedNameIntoForm();
  const onDelete = (): void => requestDeleteSelectedName();
  const onFilter = (): void => openFilterMenu();
  const onQuickCommit = (): void => commitQuickRefers();
  const onQuickCancel = (): void => cancelQuickRefers();
  const onDeleteConfirm = (): void => confirmDeleteSelectedName();
  const onDeleteCancel = (): void => closeDeleteConfirm();
  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
      return;
    }
    if (e.key === 'Enter') {
      // Enter inside an input submits the form when mutation is enabled;
      // otherwise it's an alias for Close (legacy behaviour).
      if (mayMutate()) {
        if (e.target === nameInput || e.target === formulaInput) {
          e.preventDefault();
          formRow.requestSubmit();
        }
        return;
      }
      e.preventDefault();
      api.close();
    }
  };

  shell.on(formRow, 'submit', onSubmit as EventListener);
  editorShell.on(editorCancelBtn, 'click', () => closeEditor());
  deleteShell.on(deleteOkBtn, 'click', onDeleteConfirm);
  deleteShell.on(deleteCancelBtn, 'click', onDeleteCancel);
  shell.on(newBtn, 'click', onNew);
  shell.on(editBtn, 'click', onEdit);
  shell.on(deleteBtn, 'click', onDelete);
  shell.on(filterBtn, 'click', onFilter);
  shell.on(quickCommitBtn, 'click', onQuickCommit);
  shell.on(quickCancelBtn, 'click', onQuickCancel);
  shell.on(closeBtn, 'click', onClose);
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  const api: NamedRangeDialogHandle = {
    open(entry = INSTANCE_API_NAME_MANAGER_ENTRY): void {
      dialogEntry = entry;
      applyPlatform();
      refreshFormState();
      renderList();
      shell.open();
      deferInitialFocus(overlay, () => {
        if (editorShell.isOpen() || deleteShell.isOpen()) return;
        if (mayMutate()) newBtn.focus();
        else closeBtn.focus();
      });
    },
    openNew(entry = INSTANCE_API_DEFINE_NAME_ENTRY): void {
      dialogEntry = entry;
      if (!mayMutate()) {
        api.open(entry);
        return;
      }
      applyPlatform();
      // The Mac app opens its own 名前の定義 dialog rather than the New Name editor.
      if (mac) {
        openDefine();
        return;
      }
      refreshFormState();
      renderList();
      openEditor('new');
    },
    close(): void {
      closeFilterMenu();
      closeEditor();
      closeDeleteConfirm();
      shell.close();
      host.focus();
    },
    bindWorkbook(next: WorkbookHandle): void {
      wb = next;
      refreshFormState();
    },
    detach(): void {
      closeFilterMenu();
      defineShell.dispose();
      deleteShell.dispose();
      passthroughObserver.disconnect();
      editorShell.dispose();
      shell.dispose();
    },
  };

  return api;
}
