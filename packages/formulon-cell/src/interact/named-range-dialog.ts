import { deleteDefinedName, upsertDefinedName } from '../commands/named-ranges.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';

export interface NamedRangeDialogDeps {
  host: HTMLElement;
  wb: WorkbookHandle;
  strings?: Strings;
}

export interface NamedRangeDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
  /** Swap to a fresh workbook handle (used when the host re-binds the engine). */
  bindWorkbook(next: WorkbookHandle): void;
}

/**
 * Name Manager: list of defined names plus an inline add/edit form. When the
 * engine supports `setDefinedName` (capability `definedNameMutate`), the form
 * is enabled and Add / per-row Delete write through. On engines that don't
 * (the JS stub or older bundles), the form is hidden and a note explains the
 * read-only state.
 */
export function attachNamedRangeDialog(deps: NamedRangeDialogDeps): NamedRangeDialogHandle {
  const { host } = deps;
  let wb = deps.wb;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.namedRangeDialog;

  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg fc-namedlg';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel fc-namedlg__panel';
  overlay.appendChild(panel);

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  const legendRow = document.createElement('div');
  legendRow.className = 'fc-fmtdlg__row';
  const nameHeader = document.createElement('span');
  nameHeader.textContent = t.nameHeader;
  const formulaHeader = document.createElement('span');
  formulaHeader.textContent = t.formulaHeader;
  legendRow.append(nameHeader, formulaHeader);
  body.appendChild(legendRow);

  const list = document.createElement('div');
  list.className = 'fc-namedlg__list';
  body.appendChild(list);

  // Add-row form. Only inserted into the DOM when the engine supports
  // mutation; rebuilt on bindWorkbook so capability changes are honoured.
  const formRow = document.createElement('form');
  formRow.className = 'fc-namedlg__form fc-fmtdlg__row';
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.className = 'fc-namedlg__input';
  nameInput.placeholder = t.namePlaceholder;
  nameInput.autocomplete = 'off';
  nameInput.spellcheck = false;
  const formulaInput = document.createElement('input');
  formulaInput.type = 'text';
  formulaInput.className = 'fc-namedlg__input';
  formulaInput.placeholder = t.formulaPlaceholder;
  formulaInput.autocomplete = 'off';
  formulaInput.spellcheck = false;
  const addBtn = document.createElement('button');
  addBtn.type = 'submit';
  addBtn.className = 'fc-fmtdlg__btn';
  addBtn.textContent = t.addButton;
  formRow.append(nameInput, formulaInput, addBtn);

  const errorRow = document.createElement('div');
  errorRow.className = 'fc-namedlg__error';
  errorRow.setAttribute('role', 'alert');
  errorRow.hidden = true;

  const note = document.createElement('div');
  note.className = 'fc-namedlg__note';
  note.textContent = t.note;

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);
  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-fmtdlg__btn';
  closeBtn.textContent = t.close;
  footer.appendChild(closeBtn);

  host.appendChild(overlay);

  const showError = (msg: string): void => {
    errorRow.textContent = msg;
    errorRow.hidden = false;
  };
  const clearError = (): void => {
    errorRow.hidden = true;
    errorRow.textContent = '';
  };

  const renderList = (): void => {
    list.replaceChildren();
    let count = 0;
    const canMutate = wb.capabilities.definedNameMutate;
    for (const entry of wb.definedNames()) {
      count += 1;
      const item = document.createElement('div');
      item.className = 'fc-namedlg__item fc-namedlg__row';
      const name = document.createElement('span');
      name.textContent = entry.name;
      const formulaCell = document.createElement('span');
      formulaCell.className = 'fc-namedlg__row-formula';
      const formulaText = document.createElement('span');
      formulaText.textContent = entry.formula;
      formulaCell.appendChild(formulaText);
      if (canMutate) {
        const del = document.createElement('button');
        del.type = 'button';
        del.className = 'fc-namedlg__del';
        del.textContent = t.deleteButton;
        del.addEventListener('click', () => {
          const result = deleteDefinedName(wb, entry.name);
          if (!result.ok) {
            showError(t.errorEngineFailed);
            return;
          }
          clearError();
          renderList();
        });
        formulaCell.appendChild(del);
      }
      item.append(name, formulaCell);
      list.appendChild(item);
    }
    if (count === 0) {
      const empty = document.createElement('div');
      empty.className = 'fc-namedlg__empty';
      empty.textContent = t.empty;
      list.appendChild(empty);
    }
  };

  const onSubmit = (e: SubmitEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    const name = nameInput.value.trim();
    const formula = formulaInput.value.trim();
    if (!name) {
      showError(t.errorEmptyName);
      nameInput.focus();
      return;
    }
    const result = upsertDefinedName(wb, name, formula);
    if (!result.ok) {
      showError(result.reason === 'empty-formula' ? t.errorEmptyFormula : t.errorEngineFailed);
      return;
    }
    clearError();
    nameInput.value = '';
    formulaInput.value = '';
    renderList();
    nameInput.focus();
  };

  const refreshFormState = (): void => {
    const canMutate = wb.capabilities.definedNameMutate;
    if (canMutate) {
      // Insert form before the note (which moves to the read-only fallback).
      if (!formRow.parentElement) body.append(formRow, errorRow);
      if (note.parentElement) note.remove();
    } else {
      if (formRow.parentElement) formRow.remove();
      if (errorRow.parentElement) errorRow.remove();
      if (!note.parentElement) body.appendChild(note);
    }
    clearError();
  };

  const onClose = (): void => api.close();
  const onOverlayClick = (e: MouseEvent): void => {
    if (e.target === overlay) api.close();
  };
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

  formRow.addEventListener('submit', onSubmit);
  closeBtn.addEventListener('click', onClose);
  overlay.addEventListener('click', onOverlayClick);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: NamedRangeDialogHandle = {
    open(): void {
      refreshFormState();
      renderList();
      overlay.hidden = false;
      requestAnimationFrame(() => {
        if (wb.capabilities.definedNameMutate) nameInput.focus();
        else closeBtn.focus();
      });
    },
    close(): void {
      overlay.hidden = true;
      host.focus();
    },
    bindWorkbook(next: WorkbookHandle): void {
      wb = next;
      refreshFormState();
    },
    detach(): void {
      formRow.removeEventListener('submit', onSubmit);
      closeBtn.removeEventListener('click', onClose);
      overlay.removeEventListener('click', onOverlayClick);
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };

  return api;
}
