import { writeInputValidated } from '../commands/coerce-input.js';
import { extractRefs, rotateRefAt } from '../commands/refs.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';

interface FormulaBarAutocomplete {
  isOpen(): boolean;
  move(n: number): void;
  acceptHighlighted(): boolean;
  close(): void;
  refresh(): void;
}

interface FormulaArgHelper {
  refresh(): void;
}

interface AttachFormulaBarInput {
  formulabar: HTMLElement;
  fxAccept: HTMLButtonElement;
  fxCancel: HTMLButtonElement;
  fxInput: HTMLTextAreaElement;
  getArgHelper: () => FormulaArgHelper | null;
  getAutocomplete: () => FormulaBarAutocomplete;
  cancelBindingEditor: () => void;
  host: HTMLElement;
  store: SpreadsheetStore;
  updateChrome: () => void;
  wb: () => WorkbookHandle;
}

export interface FormulaBarController {
  acceptFx(): void;
  cancelFx(): void;
  commitFx(advance: 'down' | 'right' | 'none'): void;
  detach(): void;
  isEditing(): boolean;
  refreshActions(): void;
  syncFxRefs(): void;
}

export function attachFormulaBarController(input: AttachFormulaBarInput): FormulaBarController {
  const {
    formulabar,
    fxAccept,
    fxCancel,
    fxInput,
    getArgHelper,
    getAutocomplete,
    cancelBindingEditor,
    host,
    store,
    updateChrome,
    wb,
  } = input;
  let fxEditing = false;
  let fxBaseline = '';

  const refreshActions = (): void => {
    const dirty = fxEditing && fxInput.value !== fxBaseline;
    fxCancel.disabled = !fxEditing;
    fxAccept.disabled = !dirty;
    formulabar.dataset.fcEditing = fxEditing ? '1' : '0';
  };

  const syncFxRefs = (): void => {
    const refs = extractRefs(fxInput.value).map((r) => ({
      r0: r.r0,
      c0: r.c0,
      r1: r.r1,
      c1: r.c1,
      colorIndex: r.colorIndex,
    }));
    mutators.setEditorRefs(store, refs);
  };

  const clearFxRefs = (): void => mutators.setEditorRefs(store, []);

  const commitFx = (advance: 'down' | 'right' | 'none'): void => {
    const currentWb = wb();
    const s = store.getState();
    const a = s.selection.active;
    try {
      const fmt = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
      const outcome = writeInputValidated(currentWb, a, fxInput.value, fmt?.validation);
      if (!outcome.ok) {
        console.warn(`formulon-cell: validation ${outcome.severity}: ${outcome.message}`);
        if (outcome.severity === 'stop') {
          fxInput.focus();
          return;
        }
      }
    } catch (err) {
      console.warn('formulon-cell: writeInput failed', err);
    }
    mutators.replaceCells(store, currentWb.cells(store.getState().data.sheetIndex));
    fxEditing = false;
    fxBaseline = fxInput.value;
    refreshActions();
    clearFxRefs();
    if (advance === 'down') {
      mutators.setActive(store, { ...a, row: a.row + 1 });
    } else if (advance === 'right') {
      mutators.setActive(store, { ...a, col: a.col + 1 });
    }
    host.focus();
  };

  const cancelFx = (): void => {
    fxInput.value = fxBaseline;
    fxEditing = false;
    clearFxRefs();
    getAutocomplete().close();
    refreshActions();
    host.focus();
    updateChrome();
  };

  const acceptFx = (): void => {
    if (fxInput.value !== fxBaseline) commitFx('none');
  };

  const onFxFocus = (): void => {
    cancelBindingEditor();
    fxEditing = true;
    fxBaseline = fxInput.value;
    refreshActions();
    syncFxRefs();
  };

  const onFxInput = (): void => {
    refreshActions();
    if (fxEditing) syncFxRefs();
    getAutocomplete().refresh();
    getArgHelper()?.refresh();
  };

  const onFxKeyUp = (): void => {
    if (fxEditing) getArgHelper()?.refresh();
  };

  const onFxKey = (e: KeyboardEvent): void => {
    const autocomplete = getAutocomplete();
    if (autocomplete.isOpen()) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        e.stopPropagation();
        autocomplete.move(1);
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        e.stopPropagation();
        autocomplete.move(-1);
        return;
      }
      if ((e.key === 'Enter' || e.key === 'Tab') && autocomplete.acceptHighlighted()) {
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        autocomplete.close();
        return;
      }
    }
    if (e.key === 'Enter') {
      if (e.altKey || e.shiftKey) {
        e.stopPropagation();
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      commitFx('down');
    } else if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      commitFx(e.shiftKey ? 'none' : 'right');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      fxInput.value = fxBaseline;
      fxEditing = false;
      refreshActions();
      host.focus();
      updateChrome();
    } else if (e.key === 'F4') {
      e.preventDefault();
      e.stopPropagation();
      const caret = fxInput.selectionStart ?? fxInput.value.length;
      const r = rotateRefAt(fxInput.value, caret);
      if (r.text !== fxInput.value) {
        fxInput.value = r.text;
        fxInput.setSelectionRange(r.caret, r.caret);
        syncFxRefs();
      }
    }
  };

  const onFxBlur = (): void => {
    clearFxRefs();
    getAutocomplete().close();
    if (!fxEditing) return;
    if (fxInput.value !== fxBaseline) commitFx('none');
    else {
      fxEditing = false;
      refreshActions();
    }
  };

  const keepFxFocus = (e: MouseEvent): void => e.preventDefault();

  fxInput.addEventListener('focus', onFxFocus);
  fxInput.addEventListener('input', onFxInput);
  fxInput.addEventListener('keyup', onFxKeyUp);
  fxInput.addEventListener('keydown', onFxKey);
  fxInput.addEventListener('blur', onFxBlur);
  fxCancel.addEventListener('mousedown', keepFxFocus);
  fxAccept.addEventListener('mousedown', keepFxFocus);
  fxCancel.addEventListener('click', cancelFx);
  fxAccept.addEventListener('click', acceptFx);
  refreshActions();

  return {
    acceptFx,
    cancelFx,
    commitFx,
    detach(): void {
      fxInput.removeEventListener('focus', onFxFocus);
      fxInput.removeEventListener('input', onFxInput);
      fxInput.removeEventListener('keyup', onFxKeyUp);
      fxInput.removeEventListener('keydown', onFxKey);
      fxInput.removeEventListener('blur', onFxBlur);
      fxCancel.removeEventListener('mousedown', keepFxFocus);
      fxAccept.removeEventListener('mousedown', keepFxFocus);
      fxCancel.removeEventListener('click', cancelFx);
      fxAccept.removeEventListener('click', acceptFx);
    },
    isEditing: () => fxEditing,
    refreshActions,
    syncFxRefs,
  };
}
