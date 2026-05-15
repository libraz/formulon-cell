import {
  createPivotTableFromRange,
  inferPivotSourceFields,
  type PivotSourceField,
} from '../commands/pivot-table.js';
import { PivotAggregation } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';

export interface PivotTableDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings?: Strings;
  onAfterCreate?: () => void;
  invalidate?: () => void;
}

export interface PivotTableDialogHandle {
  open(): void;
  close(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

const colLetter = (n: number): string => {
  let v = n;
  let out = '';
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const rangeLabel = (range: { r0: number; c0: number; r1: number; c1: number }): string =>
  `${colLetter(range.c0)}${range.r0 + 1}:${colLetter(range.c1)}${range.r1 + 1}`;

const parseCellRef = (input: string): { row: number; col: number } | null => {
  const m = input
    .trim()
    .replace(/\$/g, '')
    .match(/^([A-Za-z]+)([1-9][0-9]*)$/);
  if (!m) return null;
  const letters = m[1];
  const rows = m[2];
  if (!letters || !rows) return null;
  let col = 0;
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(rows) - 1, col: col - 1 };
};

const appendOption = (select: HTMLSelectElement, value: string, label: string): void => {
  const opt = document.createElement('option');
  opt.value = value;
  opt.textContent = label;
  select.appendChild(opt);
};

export function attachPivotTableDialog(deps: PivotTableDialogDeps): PivotTableDialogHandle {
  const { host, store } = deps;
  let wb = deps.wb;
  let strings = deps.strings ?? defaultStrings;
  let open = false;

  const shell = createDialogShell({
    host,
    className: 'fc-pivotdlg',
    ariaLabel: strings.pivotTableDialog.title,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-pivotdlg__panel');
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  panel.appendChild(header);

  const body = document.createElement('form');
  body.className = 'fc-fmtdlg__body fc-pivotdlg__body';
  panel.appendChild(body);

  const sourceText = document.createElement('div');
  sourceText.className = 'fc-pivotdlg__source';
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.className = 'fc-namedlg__input';
  nameInput.autocomplete = 'off';
  nameInput.spellcheck = false;
  const destInput = document.createElement('input');
  destInput.type = 'text';
  destInput.className = 'fc-namedlg__input';
  destInput.autocomplete = 'off';
  destInput.spellcheck = false;
  const rowSelect = document.createElement('select');
  rowSelect.className = 'fc-fmtdlg__select';
  const colSelect = document.createElement('select');
  colSelect.className = 'fc-fmtdlg__select';
  const valueSelect = document.createElement('select');
  valueSelect.className = 'fc-fmtdlg__select';
  const aggSelect = document.createElement('select');
  aggSelect.className = 'fc-fmtdlg__select';
  const rowSortSelect = document.createElement('select');
  rowSortSelect.className = 'fc-fmtdlg__select';
  const colSortSelect = document.createElement('select');
  colSortSelect.className = 'fc-fmtdlg__select';
  const numberFormatInput = document.createElement('input');
  numberFormatInput.type = 'text';
  numberFormatInput.className = 'fc-namedlg__input';
  numberFormatInput.autocomplete = 'off';
  numberFormatInput.spellcheck = false;
  const rowSubtotalTop = document.createElement('input');
  rowSubtotalTop.type = 'checkbox';
  rowSubtotalTop.checked = true;
  const colSubtotalTop = document.createElement('input');
  colSubtotalTop.type = 'checkbox';
  colSubtotalTop.checked = true;
  const rowTotals = document.createElement('input');
  rowTotals.type = 'checkbox';
  rowTotals.checked = true;
  const colTotals = document.createElement('input');
  colTotals.type = 'checkbox';
  colTotals.checked = true;
  const error = document.createElement('div');
  error.className = 'fc-namedlg__error';
  error.setAttribute('role', 'alert');
  error.hidden = true;

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-fmtdlg__btn';
  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
  footer.append(cancelBtn, okBtn);

  const showError = (msg: string): void => {
    error.textContent = msg;
    error.hidden = false;
  };

  const fieldSelect = (select: HTMLSelectElement, fields: readonly PivotSourceField[]): void => {
    select.replaceChildren();
    for (const f of fields) appendOption(select, f.name, f.name);
  };

  const labeled = (label: string, input: HTMLElement): HTMLLabelElement => {
    const row = document.createElement('label');
    row.className = 'fc-pivotdlg__field';
    const span = document.createElement('span');
    span.textContent = label;
    row.append(span, input);
    return row;
  };
  const checked = (label: string, input: HTMLInputElement): HTMLLabelElement => {
    const row = document.createElement('label');
    row.className = 'fc-pivotdlg__check';
    row.append(input, document.createTextNode(label));
    return row;
  };
  const section = (...children: HTMLElement[]): HTMLDivElement => {
    const el = document.createElement('div');
    el.className = 'fc-pivotdlg__section';
    el.append(...children);
    return el;
  };
  const checkGrid = (...children: HTMLLabelElement[]): HTMLDivElement => {
    const el = document.createElement('div');
    el.className = 'fc-pivotdlg__checkgrid';
    el.append(...children);
    return el;
  };

  const render = (): void => {
    const t = strings.pivotTableDialog;
    header.textContent = t.title;
    shell.setAriaLabel(t.title);
    cancelBtn.textContent = t.cancel;
    okBtn.textContent = t.ok;
    nameInput.placeholder = t.namePlaceholder;
    destInput.placeholder = t.destinationPlaceholder;
    numberFormatInput.placeholder = t.numberFormatPlaceholder;
    error.hidden = true;
    error.textContent = '';

    const range = store.getState().selection.range;
    const fields = inferPivotSourceFields(wb, range);
    const numeric = fields.filter((f) => f.numericCount > 0);
    body.replaceChildren();
    body.appendChild(sourceText);
    sourceText.textContent = `${t.source}: ${rangeLabel(range)}`;

    if (!wb.capabilities.pivotTableMutate) {
      showError(t.unsupported);
      body.appendChild(error);
      okBtn.disabled = true;
      return;
    }
    if (fields.length < 2) {
      showError(t.invalidRange);
      body.appendChild(error);
      okBtn.disabled = true;
      return;
    }

    okBtn.disabled = false;
    nameInput.value = nameInput.value || `PivotTable${wb.getPivotTables().length + 1}`;
    const dest = `${colLetter(range.c0)}${range.r1 + 3}`;
    destInput.value = destInput.value || dest;
    fieldSelect(rowSelect, fields);
    fieldSelect(colSelect, fields);
    fieldSelect(valueSelect, numeric.length > 0 ? numeric : fields);
    const noneOpt = document.createElement('option');
    noneOpt.value = '';
    noneOpt.textContent = t.none;
    colSelect.insertBefore(noneOpt, colSelect.firstChild);
    rowSelect.value = fields[0]?.name ?? '';
    valueSelect.value = (numeric[0] ?? fields[fields.length - 1])?.name ?? '';
    colSelect.value = fields[1]?.name === valueSelect.value ? '' : (fields[1]?.name ?? '');
    aggSelect.replaceChildren();
    appendOption(aggSelect, String(PivotAggregation.Sum), t.sum);
    appendOption(aggSelect, String(PivotAggregation.Count), t.count);
    rowSortSelect.replaceChildren();
    colSortSelect.replaceChildren();
    for (const select of [rowSortSelect, colSortSelect]) {
      appendOption(select, 'none', t.sortNone);
      appendOption(select, 'asc', t.sortAsc);
      appendOption(select, 'desc', t.sortDesc);
    }

    body.append(
      section(labeled(t.name, nameInput), labeled(t.destination, destInput)),
      section(
        labeled(t.rowField, rowSelect),
        labeled(t.columnField, colSelect),
        labeled(t.valueField, valueSelect),
        labeled(t.aggregation, aggSelect),
      ),
      section(
        labeled(t.rowSort, rowSortSelect),
        labeled(t.columnSort, colSortSelect),
        labeled(t.numberFormat, numberFormatInput),
      ),
      checkGrid(
        checked(t.rowSubtotalTop, rowSubtotalTop),
        checked(t.columnSubtotalTop, colSubtotalTop),
        checked(t.rowGrandTotals, rowTotals),
        checked(t.columnGrandTotals, colTotals),
      ),
      error,
    );
  };

  const close = (): void => {
    open = false;
    shell.close();
  };

  const onSubmit = (e: SubmitEvent): void => {
    e.preventDefault();
    const range = store.getState().selection.range;
    const dest = parseCellRef(destInput.value);
    if (!dest) {
      showError(strings.pivotTableDialog.invalidDestination);
      destInput.focus();
      return;
    }
    const result = createPivotTableFromRange(wb, {
      source: range,
      destination: { sheet: range.sheet, row: dest.row, col: dest.col },
      name: nameInput.value,
      rowField: rowSelect.value,
      columnField: colSelect.value || undefined,
      valueField: valueSelect.value,
      aggregation: Number(aggSelect.value) as PivotAggregation,
      rowSort: rowSortSelect.value as 'none' | 'asc' | 'desc',
      columnSort: colSortSelect.value as 'none' | 'asc' | 'desc',
      rowSubtotalTop: rowSubtotalTop.checked,
      columnSubtotalTop: colSubtotalTop.checked,
      valueNumberFormat: numberFormatInput.value,
      showRowGrandTotals: rowTotals.checked,
      showColumnGrandTotals: colTotals.checked,
    });
    if (!result.ok) {
      showError(strings.pivotTableDialog.engineFailed);
      return;
    }
    mutators.setActive(store, { sheet: range.sheet, row: dest.row, col: dest.col });
    deps.onAfterCreate?.();
    deps.invalidate?.();
    close();
  };

  const onKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    } else if (e.key === 'Enter' && !okBtn.disabled) {
      e.preventDefault();
      body.requestSubmit();
    }
  };
  const onOk = (): void => body.requestSubmit();

  shell.on(body, 'submit', onSubmit as EventListener);
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', close);
  shell.on(overlay, 'keydown', onKey as EventListener);

  return {
    open() {
      render();
      shell.open();
      open = true;
      const initial = nameInput.isConnected && !okBtn.disabled ? nameInput : cancelBtn;
      initial.focus({ preventScroll: true });
    },
    close,
    setStrings(next) {
      strings = next;
      if (open) render();
    },
    bindWorkbook(next) {
      wb = next;
      if (open) render();
    },
    detach() {
      shell.dispose();
    },
  };
}
