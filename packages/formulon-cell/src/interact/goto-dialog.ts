import {
  findMatchingCells,
  type GoToScope,
  type GoToSpecialKind,
  type GoToSpecialValueFilters,
  selectionFromMatches,
} from '../commands/goto-special.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledReason, projectDisabledState } from '../toolbar/menu-a11y.js';
import { formatA1Range } from '../wrappers/toolbar-a1.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface GoToDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so a Go To Special invocation always queries the live engine. */
  getWb: () => WorkbookHandle;
  strings?: Strings;
}

export interface GoToDialogHandle {
  open(mode?: 'go-to' | 'special'): void;
  close(): void;
  detach(): void;
}

const KINDS: readonly GoToSpecialKind[] = [
  'blanks',
  'non-blanks',
  'formulas',
  'constants',
  'numbers',
  'text',
  'errors',
  'data-validation',
  'conditional-format',
];

/**
 * Spreadsheet-style "Go To Special" dialog. Lets the user pick a category
 * (blanks, formulas, errors, validation, …) and rewrites the selection to the
 * bounding range of every matching cell on the active sheet (or just inside
 * the current selection when one is provided).
 *
 * Lifecycle mirrors the other modals: `attach…()` mounts a hidden overlay on
 * the host; `open()` shows it, `close()` hides it, `detach()` tears it down.
 * The OK button runs the predicate; on no matches an inline status message is
 * shown and the dialog stays open. On 1+ matches the selection jumps to the
 * bounding rect and the dialog closes.
 */
export function attachGoToDialog(deps: GoToDialogDeps): GoToDialogHandle {
  const { host, store, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.goToDialog;
  let mode: 'go-to' | 'special' = 'special';

  const shell = createDialogShell({
    host,
    className: 'fc-goto',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  // Reuse the shared format-dialog skin for header/footer/btn styling.
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-goto__panel'],
  });

  // ── Direct reference (normal Go To) ───────────────────────────────────
  const referenceRow = document.createElement('label');
  referenceRow.className = 'fc-goto__reference';
  const referenceLabel = document.createElement('span');
  referenceLabel.textContent = t.reference;
  const referenceInput = document.createElement('input');
  referenceInput.type = 'text';
  referenceInput.placeholder = t.referencePlaceholder;
  referenceInput.autocomplete = 'off';
  referenceInput.spellcheck = false;
  referenceRow.append(referenceLabel, referenceInput);
  attachRangePickerButton(referenceInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => formatA1Range(store.getState().selection.range),
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'go-to-reference',
  });
  body.appendChild(referenceRow);

  // ── Scope (only meaningful when current selection is multi-cell) ───────
  const scopeLegend = document.createElement('div');
  scopeLegend.className = 'fc-goto__legend';
  scopeLegend.textContent = t.scopeLabel;
  body.appendChild(scopeLegend);

  const scopeGroup = document.createElement('div');
  scopeGroup.className = 'fc-goto__scope';
  scopeGroup.setAttribute('role', 'radiogroup');
  scopeGroup.setAttribute('aria-label', t.scopeLabel);
  body.appendChild(scopeGroup);

  const scopeName = `fc-goto-scope-${Math.random().toString(36).slice(2, 8)}`;
  const makeScopeRadio = (value: GoToScope, label: string): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-goto__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = scopeName;
    input.value = value;
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    scopeGroup.appendChild(wrap);
    return input;
  };
  const scopeSheet = makeScopeRadio('sheet', t.scopeSheet);
  const scopeSelection = makeScopeRadio('selection', t.scopeSelection);
  scopeSheet.checked = true;

  // ── Kind list ──────────────────────────────────────────────────────────
  const kindLegend = document.createElement('div');
  kindLegend.className = 'fc-goto__legend';
  kindLegend.textContent = t.kindLabel;
  body.appendChild(kindLegend);

  const kindList = document.createElement('div');
  kindList.className = 'fc-goto__kinds';
  kindList.setAttribute('role', 'radiogroup');
  kindList.setAttribute('aria-label', t.kindLabel);
  body.appendChild(kindList);

  const kindName = `fc-goto-kind-${Math.random().toString(36).slice(2, 8)}`;
  const kindLabels: Record<GoToSpecialKind, string> = {
    blanks: t.kindBlanks,
    'non-blanks': t.kindNonBlanks,
    formulas: t.kindFormulas,
    constants: t.kindConstants,
    numbers: t.kindNumbers,
    text: t.kindText,
    errors: t.kindErrors,
    'data-validation': t.kindDataValidation,
    'conditional-format': t.kindConditionalFormat,
  };
  const kindInputs = new Map<GoToSpecialKind, HTMLInputElement>();
  for (const k of KINDS) {
    const wrap = document.createElement('label');
    wrap.className = 'fc-goto__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = kindName;
    input.value = k;
    const span = document.createElement('span');
    span.textContent = kindLabels[k];
    wrap.append(input, span);
    kindList.appendChild(wrap);
    kindInputs.set(k, input);
  }
  // Default selection — spreadsheets open with `constants` highlighted; we mirror that.
  (kindInputs.get('constants') as HTMLInputElement).checked = true;

  // Excel shows value-type checkboxes when either Formulas or Constants is
  // selected. They let the user narrow the match to number/text/logical/error
  // results without changing the top-level category.
  const valueFilters = document.createElement('fieldset');
  valueFilters.className = 'fc-goto__value-filters';
  const valueLegend = document.createElement('legend');
  valueLegend.textContent = t.valueFilterLabel;
  valueFilters.appendChild(valueLegend);
  body.appendChild(valueFilters);

  const makeValueFilter = (key: keyof GoToSpecialValueFilters, label: string): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-goto__check';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.value = key;
    input.checked = true;
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    valueFilters.appendChild(wrap);
    return input;
  };
  const valueFilterInputs = {
    numbers: makeValueFilter('numbers', t.kindNumbers),
    text: makeValueFilter('text', t.kindText),
    logical: makeValueFilter('logical', t.kindLogical),
    errors: makeValueFilter('errors', t.kindErrors),
  };

  // Inline status (shown when a search returns zero results).
  const statusLine = document.createElement('div');
  statusLine.className = 'fc-goto__status';
  statusLine.setAttribute('role', 'status');
  statusLine.setAttribute('aria-live', 'polite');
  body.appendChild(statusLine);

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
  });

  const isSelectionMulti = (): boolean => {
    const r = store.getState().selection.range;
    return r.r1 > r.r0 || r.c1 > r.c0;
  };

  const syncScopeAvailability = (): void => {
    const multi = isSelectionMulti();
    // Disable selection scope when only one cell is selected — matching the
    // spreadsheet convention, which silently widens to "active sheet" in
    // that case.
    const reason = multi ? null : t.scopeSelectionRequiresMultiCell;
    projectDisabledState(scopeSelection, !multi, reason, { datasetKey: 'disabledReason' });
    const scopeSelectionLabel = scopeSelection.closest<HTMLElement>('label');
    if (scopeSelectionLabel) {
      projectDisabledReason(scopeSelectionLabel, reason, {
        ariaDescription: false,
        titlePrefix: t.scopeSelection,
      });
    }
    if (!multi) {
      scopeSheet.checked = true;
    }
  };

  const getCheckedKind = (): GoToSpecialKind => {
    for (const [k, input] of kindInputs) if (input.checked) return k;
    return 'constants';
  };
  const getValueFilters = (kind: GoToSpecialKind): GoToSpecialValueFilters | undefined => {
    if (kind !== 'formulas' && kind !== 'constants') return undefined;
    return {
      numbers: valueFilterInputs.numbers.checked,
      text: valueFilterInputs.text.checked,
      logical: valueFilterInputs.logical.checked,
      errors: valueFilterInputs.errors.checked,
    };
  };
  const getCheckedScope = (): GoToScope => (scopeSelection.checked ? 'selection' : 'sheet');
  const sheetIndexByName = (name: string): number => {
    const target = name.toLowerCase();
    const wb = getWb();
    for (let i = 0; i < wb.sheetCount; i += 1) {
      if (wb.sheetName(i).toLowerCase() === target) return i;
    }
    return -1;
  };

  const goToReference = (): boolean => {
    const parsed = parseRangeRef(referenceInput.value);
    if (!parsed) {
      statusLine.textContent = t.invalidReference;
      return false;
    }
    const currentSheet = store.getState().data.sheetIndex;
    const sheet = parsed.sheetName == null ? currentSheet : sheetIndexByName(parsed.sheetName);
    if (sheet < 0) {
      statusLine.textContent = t.invalidReference;
      return false;
    }
    const active = { sheet, row: parsed.r0, col: parsed.c0 };
    store.setState((s) => ({
      ...s,
      data: { ...s.data, sheetIndex: sheet },
      selection: {
        active,
        anchor: active,
        range: { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 },
        extraRanges: [],
      },
    }));
    api.close();
    return true;
  };

  const onOk = (): void => {
    statusLine.textContent = '';
    if (mode === 'go-to') {
      goToReference();
      return;
    }
    const kind = getCheckedKind();
    const scope = getCheckedScope();
    const wb = getWb();
    const matches = findMatchingCells(wb, store, scope, kind, getValueFilters(kind));
    if (matches.length === 0) {
      statusLine.textContent = t.noResults;
      return;
    }
    store.setState((s) => ({
      ...s,
      selection: selectionFromMatches(matches),
    }));
    api.close();
  };

  const syncValueFilters = (): void => {
    const kind = getCheckedKind();
    valueFilters.hidden = mode === 'go-to' || (kind !== 'formulas' && kind !== 'constants');
  };

  for (const input of kindInputs.values()) {
    shell.on(input, 'change', syncValueFilters);
  }

  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', () => api.close());
  // Enter on any control inside the modal commits — matches the spreadsheet
  // convention. Escape and backdrop dismissals are handled by the shell.
  shell.on(shell.overlay, 'keydown', (e) => {
    if ((e as KeyboardEvent).key === 'Enter') {
      e.preventDefault();
      onOk();
    }
  });

  const api: GoToDialogHandle = {
    open(nextMode = 'special'): void {
      mode = nextMode;
      statusLine.textContent = '';
      const normal = mode === 'go-to';
      header.textContent = normal ? t.goToTitle : t.title;
      shell.setAriaLabel(normal ? t.goToTitle : t.title);
      referenceRow.hidden = !normal;
      scopeLegend.hidden = normal;
      scopeGroup.hidden = normal;
      kindLegend.hidden = normal;
      kindList.hidden = normal;
      syncValueFilters();
      referenceInput.value = normal ? '' : referenceInput.value;
      syncScopeAvailability();
      if (!normal && isSelectionMulti()) scopeSelection.checked = true;
      shell.open();
      requestAnimationFrame(() => {
        if (normal) {
          referenceInput.focus();
          return;
        }
        const first = kindInputs.values().next().value;
        if (first) first.focus();
      });
    },
    close(): void {
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };
  return api;
}
