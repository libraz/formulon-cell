import type { History } from '../commands/history.js';
import {
  deleteDefinedName,
  recordDefinedNamesChange,
  upsertDefinedName,
} from '../commands/named-ranges.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogButton,
  appendDialogIconButton,
  clearDialogError,
  createDialogButton,
  createDialogShell,
  focusAndSelectInput,
  showDialogError,
} from './dialog-shell.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface NamedRangeDialogDeps {
  host: HTMLElement;
  wb: WorkbookHandle;
  history?: History | null;
  strings?: Strings;
  getSelectedRangeFormula?: () => string;
  subscribeToRangeChanges?: (listener: () => void) => () => void;
  /** Called after a defined-name mutation writes through to the engine so the
   *  host can re-project recalculated cells into the store (H-38). */
  onAfterMutate?: () => void;
}

export interface NamedRangeDialogHandle {
  open(): void;
  openNew(): void;
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
): string {
  if (reason === 'invalid-name') return t.errorInvalidName;
  if (reason === 'empty-name') return t.errorEmptyName;
  if (reason === 'empty-formula') return t.errorEmptyFormula;
  return t.errorEngineFailed;
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
  const strings = deps.strings ?? defaultStrings;
  const t = strings.namedRangeDialog;
  const rangePickerOptions =
    deps.getSelectedRangeFormula && deps.subscribeToRangeChanges
      ? {
          label: strings.pivotTableDialog.rangePickerSelect,
          getValue: deps.getSelectedRangeFormula,
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
  for (const { key, label } of headerSpecs) {
    const cell = document.createElement('span');
    const button = createNameManagerButton(label, 'fc-namedlg__sort');
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
  quickLabel.textContent = t.formulaHeader;
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
    label: '✓',
    ariaLabel: t.commitButton,
    baseClass: 'fc-namedlg__refers-icon',
  });
  const quickCancelBtn = appendDialogIconButton(quickRow, {
    label: '×',
    ariaLabel: t.cancelButton,
    baseClass: 'fc-namedlg__refers-icon',
  });
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
  const scopeSelect = createDialogSelect(
    [{ value: 'workbook', label: t.workbookScope }],
    'workbook',
    { className: 'fc-namedlg-editor__input', ariaLabel: t.scopeHeader },
  );
  projectDisabledState(scopeSelect, true, t.scopeWorkbookOnly, {
    datasetKey: 'disabledReason',
    titlePrefix: t.scopeHeader,
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
  formRow.append(
    makeEditorRow(t.nameHeader, nameInput),
    makeEditorRow(t.scopeHeader, scopeSelect),
    makeEditorRow(t.commentHeader, commentInput),
    makeEditorRow(t.formulaHeader, formulaInput),
  );
  if (rangePickerOptions) {
    attachRangePickerButton(formulaInput, {
      ...rangePickerOptions,
      kind: 'named-range-editor-refers-to',
    });
  }
  const editorButtons = document.createElement('div');
  editorButtons.className = 'fc-namedlg-editor__buttons';
  const editorOkBtn = appendDialogButton(editorButtons, { label: t.ok, variant: 'primary' });
  editorOkBtn.type = 'submit';
  const editorCancelBtn = appendDialogButton(editorButtons, { label: t.cancel });
  formRow.appendChild(editorButtons);
  const errorRow = document.createElement('div');
  errorRow.className = 'fc-namedlg__error';
  errorRow.setAttribute('role', 'alert');
  errorRow.hidden = true;
  editorBody.append(formRow, errorRow);

  const deleteShell = createDialogShell({
    host,
    className: 'fc-namedlg-confirm',
    ariaLabel: t.confirmDeleteTitle,
    onDismiss: () => closeDeleteConfirm(),
  });
  deleteShell.overlay.classList.add('fc-fmtdlg');
  deleteShell.panel.classList.add('fc-fmtdlg__panel', 'fc-namedlg-confirm__panel');
  const deleteTitle = document.createElement('div');
  deleteTitle.className = 'fc-fmtdlg__header';
  deleteTitle.textContent = t.confirmDeleteTitle;
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
  let currentRows: { name: string; formula: string }[] = [];
  let activeFilter: NameFilter = 'all';
  let sortKey: NameSortKey = 'name';
  let sortDir: 'asc' | 'desc' = 'asc';
  let filterMenu: HTMLDivElement | null = null;
  let editorMode: 'new' | 'edit' = 'new';
  let pendingDeleteName: string | null = null;

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
    const canMutate = wb.capabilities.definedNameMutate;
    const enabled = canMutate && Boolean(entry);
    quickInput.value = entry?.formula ?? '';
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
    const canMutate = wb.capabilities.definedNameMutate;
    const allRows = [...wb.definedNames()];
    currentRows = allRows.filter((entry) => {
      if (activeFilter === 'errors') return hasFormulaError(entry.formula);
      if (activeFilter === 'noErrors') return !hasFormulaError(entry.formula);
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
      value.textContent = t.valueUnavailable;
      const formulaCell = document.createElement('span');
      formulaCell.className = 'fc-namedlg__row-formula';
      formulaCell.textContent = entry.formula;
      const scope = document.createElement('span');
      scope.className = 'fc-namedlg__row-scope';
      scope.textContent = t.workbookScope;
      const comment = document.createElement('span');
      comment.className = 'fc-namedlg__row-comment';
      comment.textContent = '';
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
          deleteBtn.click();
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
      empty.textContent = t.empty;
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

  const sortValue = (entry: { name: string; formula: string }, key: NameSortKey): string => {
    switch (key) {
      case 'formula':
        return entry.formula;
      case 'scope':
        return t.workbookScope;
      case 'value':
        return t.valueUnavailable;
      case 'comment':
        return '';
      default:
        return entry.name;
    }
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
    document.removeEventListener('pointerdown', onFilterDocPointer, true);
    document.removeEventListener('keydown', onFilterDocKey, true);
  };

  function onFilterDocPointer(e: PointerEvent): void {
    const target = e.target;
    if (
      target instanceof Node &&
      (filterMenu?.contains(target) === true || filterBtn.contains(target))
    ) {
      return;
    }
    closeFilterMenu();
  }

  function onFilterDocKey(e: KeyboardEvent): void {
    if (!filterMenu) return;
    const items = Array.from(filterMenu.querySelectorAll<HTMLButtonElement>('[role="menuitem"]'));
    const active =
      document.activeElement instanceof HTMLButtonElement ? document.activeElement : null;
    const idx = active ? items.indexOf(active) : -1;
    const focusAt = (next: number): void => {
      e.preventDefault();
      e.stopPropagation();
      const wrapped = (next + items.length) % items.length;
      items[wrapped]?.focus();
    };
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      closeFilterMenu();
      filterBtn.focus();
    } else if (e.key === 'ArrowDown') {
      focusAt(idx < 0 ? 0 : idx + 1);
    } else if (e.key === 'ArrowUp') {
      focusAt(idx < 0 ? items.length - 1 : idx - 1);
    } else if (e.key === 'Home') {
      focusAt(0);
    } else if (e.key === 'End') {
      focusAt(items.length - 1);
    }
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
    document.body.appendChild(menu);
    const r = filterBtn.getBoundingClientRect();
    menu.style.left = `${Math.max(4, r.left)}px`;
    menu.style.top = `${r.bottom + 2}px`;
    filterMenu = menu;
    document.addEventListener('pointerdown', onFilterDocPointer, true);
    document.addEventListener('keydown', onFilterDocKey, true);
    menu.querySelector<HTMLButtonElement>('[role="menuitem"]')?.focus();
  };

  function closeEditor(): void {
    editorShell.overlay.dispatchEvent(new CustomEvent('fc-range-picker-stop-all'));
    editorShell.close();
    clearError();
    if (shell.isOpen()) panel.focus();
  }

  const openEditor = (mode: 'new' | 'edit'): void => {
    const entry = mode === 'edit' ? currentRows[selectedNameIndex] : null;
    if (mode === 'edit' && !entry) return;
    overlay.dispatchEvent(new CustomEvent('fc-range-picker-stop-all'));
    editorMode = mode;
    const title = mode === 'new' ? t.newNameTitle : t.editNameTitle;
    editorTitle.textContent = title;
    editorShell.setAriaLabel(title);
    nameInput.value = entry?.name ?? '';
    scopeSelect.value = 'workbook';
    commentInput.value = '';
    formulaInput.value = entry?.formula ?? '';
    clearError();
    editorShell.open();
    requestAnimationFrame(() => {
      focusAndSelectInput(nameInput);
    });
  };

  const loadSelectedNameIntoForm = (): void => {
    const entry = currentRows[selectedNameIndex];
    if (!entry) return;
    openEditor('edit');
  };

  function closeDeleteConfirm(): void {
    deleteShell.close();
    pendingDeleteName = null;
    if (shell.isOpen()) panel.focus();
  }

  const requestDeleteSelectedName = (): void => {
    const entry = currentRows[selectedNameIndex];
    if (!entry) return;
    pendingDeleteName = entry.name;
    deleteMessage.textContent = t.confirmDeleteMessage.replace('{name}', entry.name);
    deleteShell.open();
    requestAnimationFrame(() => deleteOkBtn.focus());
  };

  const confirmDeleteSelectedName = (): void => {
    if (!pendingDeleteName) return;
    const name = pendingDeleteName;
    const result = recordDefinedNamesChange(history, wb, () => deleteDefinedName(wb, name));
    if (!result.ok) {
      showMainError(t.errorEngineFailed);
      closeDeleteConfirm();
      return;
    }
    clearMainError();
    deps.onAfterMutate?.();
    closeDeleteConfirm();
    selectedNameIndex = Math.min(selectedNameIndex, Math.max(0, currentRows.length - 2));
    renderList();
    requestAnimationFrame(() => focusNameRow(selectedNameIndex));
  };

  const commitQuickRefers = (): void => {
    const entry = currentRows[selectedNameIndex];
    if (!entry) return;
    const result = recordDefinedNamesChange(history, wb, () =>
      upsertDefinedName(wb, entry.name, quickInput.value),
    );
    if (!result.ok) {
      showMainError(errorMessageFor(result.reason, t));
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
    const result = recordDefinedNamesChange(history, wb, () =>
      upsertDefinedName(wb, name, formula),
    );
    if (!result.ok) {
      showError(errorMessageFor(result.reason, t));
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
    if (editorMode === 'new') newBtn.focus();
    else focusNameRow(selectedNameIndex);
  };

  const refreshFormState = (): void => {
    const canMutate = wb.capabilities.definedNameMutate;
    if (!canMutate) {
      setReadOnlyState(filterBtn, currentRows.length === 0);
    }
    setReadOnlyState(newBtn, !canMutate);
    setReadOnlyState(editBtn, !canMutate);
    setReadOnlyState(deleteBtn, !canMutate);
    setReadOnlyState(quickInput, !canMutate);
    setReadOnlyState(quickCommitBtn, !canMutate);
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
      if (wb.capabilities.definedNameMutate) {
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
    open(): void {
      refreshFormState();
      renderList();
      shell.open();
      requestAnimationFrame(() => {
        if (wb.capabilities.definedNameMutate) nameInput.focus();
        else closeBtn.focus();
      });
    },
    openNew(): void {
      api.open();
      if (!wb.capabilities.definedNameMutate) return;
      requestAnimationFrame(() => openEditor('new'));
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
      deleteShell.dispose();
      editorShell.dispose();
      shell.dispose();
    },
  };

  return api;
}
