import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SlicerSpec, SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';

export interface MacSlicerDialogInput {
  tableName: string;
  column: string;
}

export interface MacSlicerDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle;
  strings?: Strings;
  /** Called after the user picks a table and field. Returning `null` leaves
   *  the dialog open with an error, which lets the public instance API own
   *  feature/policy validation. */
  onAdd: (input: MacSlicerDialogInput) => SlicerSpec | null;
}

export interface MacSlicerDialogHandle {
  open(): void;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  detach(): void;
}

interface TableOption {
  name: string;
  displayName: string;
  columns: readonly string[];
}

/** Insert Slicer chooser. The existing slicer feature renders the
 *  resulting panel; this dialog only resolves a real workbook table + column
 *  and delegates creation to the instance action provider. */
export function attachMacSlicerDialog(deps: MacSlicerDialogDeps): MacSlicerDialogHandle {
  const { host } = deps;
  let strings = deps.strings ?? defaultStrings;
  let tables: TableOption[] = [];

  const shell = createDialogShell({
    host,
    className: 'fc-macslicerdlg',
    ariaLabel: strings.slicer.addSlicer,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: strings.slicer.addSlicer,
    panelClasses: ['fc-fmtdlg__panel', 'fc-macslicerdlg__panel'],
    bodyClass: 'fc-fmtdlg__body fc-macslicerdlg__body',
  });

  const tableSelect = document.createElement('select');
  tableSelect.className = 'fc-fmtdlg__select fc-macslicerdlg__table';
  tableSelect.dataset.macSlicerTable = 'true';
  tableSelect.setAttribute('aria-label', strings.macSlicer.table);
  const columnSelect = document.createElement('select');
  columnSelect.className = 'fc-fmtdlg__select fc-macslicerdlg__column';
  columnSelect.dataset.macSlicerColumn = 'true';
  columnSelect.setAttribute('aria-label', strings.slicer.chooseColumn);
  const error = document.createElement('div');
  error.className = 'fc-fmtdlg__error fc-macslicerdlg__error';
  error.dataset.macSlicerError = 'true';
  error.setAttribute('role', 'alert');
  error.hidden = true;

  const row = (control: HTMLElement): { wrapper: HTMLLabelElement; text: HTMLSpanElement } => {
    const wrapper = document.createElement('label');
    wrapper.className = 'fc-fmtdlg__row fc-macslicerdlg__row';
    const text = document.createElement('span');
    wrapper.append(text, control);
    return { wrapper, text };
  };
  const tableRow = row(tableSelect);
  const columnRow = row(columnSelect);
  body.append(tableRow.wrapper, columnRow.wrapper, error);
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: strings.macSlicer.cancel,
    okLabel: strings.macSlicer.ok,
  });

  const setError = (message: string | null): void => {
    error.hidden = message === null;
    error.textContent = message ?? '';
    projectDisabledState(okBtn, message !== null, message, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.macSlicer.ok,
    });
  };

  const renderColumns = (): void => {
    const table = tables.find((entry) => entry.name === tableSelect.value);
    const previous = columnSelect.value;
    columnSelect.replaceChildren();
    if (!table || table.columns.length === 0) {
      setError(tables.length === 0 ? strings.macSlicer.noTable : strings.macSlicer.noColumns);
      return;
    }
    appendDialogSelectOptions(
      columnSelect,
      table.columns.map((column) => ({ value: column, label: column })),
    );
    columnSelect.value = table.columns.includes(previous) ? previous : (table.columns[0] ?? '');
    setError(null);
  };

  const renderTables = (): void => {
    const previous = tableSelect.value;
    tables = deps
      .getWb()
      .getTables()
      .map((table) => ({
        name: table.name,
        displayName: table.displayName || table.name,
        columns: [...table.columns],
      }));
    tableSelect.replaceChildren();
    appendDialogSelectOptions(
      tableSelect,
      tables.map((table) => ({ value: table.name, label: table.displayName })),
    );
    if (tables.length > 0) {
      tableSelect.value = tables.some((table) => table.name === previous)
        ? previous
        : (tables[0]?.name ?? '');
    }
    renderColumns();
  };

  const onTableChange = (): void => renderColumns();
  const onOk = (): void => {
    const table = tables.find((entry) => entry.name === tableSelect.value);
    const column = columnSelect.value;
    if (!table || !column) {
      setError(strings.macSlicer.chooseTableAndColumn);
      return;
    }
    const result = deps.onAdd({ tableName: table.name, column });
    if (!result) {
      setError(strings.macSlicer.insertFailed);
      return;
    }
    api.close();
  };
  const onCancel = (): void => api.close();

  shell.on(tableSelect, 'change', onTableChange);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(okBtn, 'click', onOk);

  const refreshLabels = (): void => {
    header.textContent = strings.slicer.addSlicer;
    shell.setAriaLabel(strings.slicer.addSlicer);
    tableRow.text.textContent = strings.macSlicer.table;
    columnRow.text.textContent = strings.slicer.chooseColumn;
    tableSelect.setAttribute('aria-label', strings.macSlicer.table);
    columnSelect.setAttribute('aria-label', strings.slicer.chooseColumn);
    cancelBtn.textContent = strings.macSlicer.cancel;
    okBtn.textContent = strings.macSlicer.ok;
  };

  const api: MacSlicerDialogHandle = {
    open(): void {
      renderTables();
      shell.open();
      requestAnimationFrame(() => tableSelect.focus());
    },
    close(): void {
      shell.close();
      host.focus();
    },
    refresh(): void {
      refreshLabels();
      if (shell.isOpen()) renderTables();
    },
    setStrings(next: Strings): void {
      strings = next;
      refreshLabels();
      if (shell.isOpen()) renderTables();
    },
    detach(): void {
      shell.dispose();
    },
  };

  refreshLabels();
  return api;
}
