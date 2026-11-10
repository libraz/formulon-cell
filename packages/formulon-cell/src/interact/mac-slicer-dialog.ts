import { activeTableOverlay } from '../commands/table-design.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SlicerSpec, SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';

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

  // Excel lists every field of the table; each checked one becomes a slicer.
  const fieldList = document.createElement('div');
  fieldList.className = 'fc-macslicerdlg__fields';
  fieldList.setAttribute('role', 'group');
  const error = document.createElement('div');
  error.className = 'fc-fmtdlg__error fc-macslicerdlg__error';
  error.dataset.macSlicerError = 'true';
  error.setAttribute('role', 'alert');
  error.hidden = true;
  body.append(fieldList, error);
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: strings.mac.slicer.cancel,
    okLabel: strings.mac.slicer.ok,
  });

  const setError = (message: string | null): void => {
    error.hidden = message === null;
    error.textContent = message ?? '';
    projectDisabledState(okBtn, message !== null, message, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.mac.slicer.ok,
    });
  };

  /** The table under the active cell, else the first one. */
  const targetTable = (): TableOption | undefined => {
    const active = activeTableOverlay(deps.store.getState());
    return tables.find((entry) => entry.name === active?.name) ?? tables[0];
  };

  const renderFields = (): void => {
    tables = deps
      .getWb()
      .getTables()
      .map((table) => ({
        name: table.name,
        displayName: table.displayName || table.name,
        columns: [...table.columns],
      }));
    fieldList.replaceChildren();
    const table = targetTable();
    if (!table || table.columns.length === 0) {
      setError(tables.length === 0 ? strings.mac.slicer.noTable : strings.mac.slicer.noColumns);
      return;
    }
    for (const column of table.columns) {
      const row = document.createElement('label');
      row.className = 'fc-macslicerdlg__field';
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.value = column;
      const text = document.createElement('span');
      text.textContent = column;
      row.append(box, text);
      fieldList.appendChild(row);
    }
    setError(null);
  };

  const onOk = (): void => {
    const table = targetTable();
    const checked = Array.from(
      fieldList.querySelectorAll<HTMLInputElement>('input[type="checkbox"]:checked'),
      (box) => box.value,
    );
    if (!table || checked.length === 0) {
      api.close();
      return;
    }
    for (const column of checked) {
      if (!deps.onAdd({ tableName: table.name, column })) {
        setError(strings.mac.slicer.insertFailed);
        return;
      }
    }
    api.close();
  };
  const onCancel = (): void => api.close();

  shell.on(cancelBtn, 'click', onCancel);
  shell.on(okBtn, 'click', onOk);

  const refreshLabels = (): void => {
    header.textContent = strings.slicer.addSlicer;
    shell.setAriaLabel(strings.slicer.addSlicer);
    fieldList.setAttribute('aria-label', strings.slicer.chooseColumn);
    cancelBtn.textContent = strings.mac.slicer.cancel;
    okBtn.textContent = strings.mac.slicer.ok;
  };

  const api: MacSlicerDialogHandle = {
    open(): void {
      renderFields();
      shell.open();
      deferInitialFocus(shell.overlay, () => fieldList.querySelector('input')?.focus());
    },
    close(): void {
      shell.close();
      host.focus();
    },
    refresh(): void {
      refreshLabels();
      if (shell.isOpen()) renderFields();
    },
    setStrings(next: Strings): void {
      strings = next;
      refreshLabels();
      if (shell.isOpen()) renderFields();
    },
    detach(): void {
      shell.dispose();
    },
  };

  refreshLabels();
  return api;
}
