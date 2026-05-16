import {
  appendDialogActions,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';
import type { SortDialogColumn } from './sort.js';

export interface RemoveDuplicatesDialogOptions {
  title: string;
  columnsLabel: string;
  headerLabel: string;
  selectAllLabel: string;
  unselectAllLabel: string;
  noColumnsLabel: string;
  columns: readonly SortDialogColumn[];
  initialColumns: readonly string[];
  initialHasHeader: boolean;
  okLabel?: string;
  cancelLabel?: string;
}

export interface RemoveDuplicatesDialogResult {
  columns: string[];
  hasHeader: boolean;
}

export const showRemoveDuplicatesDialog = (
  opts: RemoveDuplicatesDialogOptions,
): Promise<RemoveDuplicatesDialogResult | null> =>
  new Promise<RemoveDuplicatesDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });

    const headerRow = document.createElement('label');
    headerRow.className = 'fc-fmtdlg__row app__dlg__label';
    const hasHeader = document.createElement('input');
    hasHeader.type = 'checkbox';
    hasHeader.checked = opts.initialHasHeader;
    headerRow.append(hasHeader, document.createTextNode(` ${opts.headerLabel}`));
    shell.body.appendChild(headerRow);

    const actionRow = document.createElement('div');
    actionRow.className = 'fc-fmtdlg__row';
    const selectAllBtn = document.createElement('button');
    selectAllBtn.type = 'button';
    selectAllBtn.className = 'fc-fmtdlg__btn';
    selectAllBtn.textContent = opts.selectAllLabel;
    const unselectAllBtn = document.createElement('button');
    unselectAllBtn.type = 'button';
    unselectAllBtn.className = 'fc-fmtdlg__btn';
    unselectAllBtn.textContent = opts.unselectAllLabel;
    actionRow.append(selectAllBtn, unselectAllBtn);
    shell.body.appendChild(actionRow);

    const fieldset = document.createElement('fieldset');
    fieldset.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
    const legend = document.createElement('legend');
    legend.className = 'app__dlg__label';
    legend.textContent = opts.columnsLabel;
    fieldset.appendChild(legend);
    const checks: HTMLInputElement[] = [];
    const initialColumns = new Set(opts.initialColumns);
    for (const item of opts.columns) {
      const label = document.createElement('label');
      label.className = 'app__dlg__label';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = item.value;
      checkbox.checked = initialColumns.has(item.value);
      checks.push(checkbox);
      label.append(checkbox, document.createTextNode(` ${item.label}`));
      fieldset.appendChild(label);
    }
    shell.body.appendChild(fieldset);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel ?? 'Cancel',
      okLabel: opts.okLabel ?? 'OK',
    });

    const lifecycle = installDialogLifecycle<RemoveDuplicatesDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const selected = checks
        .filter((checkbox) => checkbox.checked)
        .map((checkbox) => checkbox.value);
      if (selected.length === 0) {
        errorRow.textContent = opts.noColumnsLabel;
        errorRow.hidden = false;
        return;
      }
      lifecycle.finish({ columns: selected, hasHeader: hasHeader.checked });
    };
    selectAllBtn.addEventListener('click', () => {
      for (const checkbox of checks) checkbox.checked = true;
      errorRow.hidden = true;
    });
    unselectAllBtn.addEventListener('click', () => {
      for (const checkbox of checks) checkbox.checked = false;
    });
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, checks[0] ?? okBtn);
  });
