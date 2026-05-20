import {
  appendDialogActions,
  appendDialogButton,
  appendErrorRow,
  clearDialogError,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
  showDialogError,
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
  okLabel: string;
  cancelLabel: string;
}

export interface RemoveDuplicatesDialogResult {
  columns: string[];
  hasHeader: boolean;
}

export const showRemoveDuplicatesDialog = (
  opts: RemoveDuplicatesDialogOptions,
): Promise<RemoveDuplicatesDialogResult | null> =>
  new Promise<RemoveDuplicatesDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });

    const headerRow = document.createElement('label');
    headerRow.className = 'fc-fmtdlg__row app__dlg__label fc-dedupedlg__header';
    const hasHeader = document.createElement('input');
    hasHeader.type = 'checkbox';
    hasHeader.checked = opts.initialHasHeader;
    headerRow.append(hasHeader, document.createTextNode(` ${opts.headerLabel}`));
    shell.body.appendChild(headerRow);

    const actionRow = document.createElement('div');
    actionRow.className = 'fc-dedupedlg__actions';
    const selectAllBtn = appendDialogButton(actionRow, { label: opts.selectAllLabel });
    const unselectAllBtn = appendDialogButton(actionRow, { label: opts.unselectAllLabel });
    shell.body.appendChild(actionRow);

    const fieldset = document.createElement('fieldset');
    fieldset.className = 'fc-dedupedlg__columns';
    const legend = document.createElement('legend');
    legend.className = 'fc-dedupedlg__legend';
    legend.textContent = opts.columnsLabel;
    fieldset.appendChild(legend);
    const list = document.createElement('div');
    list.className = 'fc-dedupedlg__column-list';
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', opts.columnsLabel);
    fieldset.appendChild(list);
    const checks: HTMLInputElement[] = [];
    const initialColumns = new Set(opts.initialColumns);
    for (const item of opts.columns) {
      const label = document.createElement('label');
      label.className = 'fc-dedupedlg__column';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = item.value;
      checkbox.checked = initialColumns.has(item.value);
      checks.push(checkbox);
      const text = document.createElement('span');
      text.textContent = item.label;
      label.append(checkbox, text);
      list.appendChild(label);
    }
    shell.body.appendChild(fieldset);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
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
        showDialogError(errorRow, opts.noColumnsLabel);
        return;
      }
      lifecycle.finish({ columns: selected, hasHeader: hasHeader.checked });
    };
    selectAllBtn.addEventListener('click', () => {
      for (const checkbox of checks) checkbox.checked = true;
      clearDialogError(errorRow);
    });
    unselectAllBtn.addEventListener('click', () => {
      for (const checkbox of checks) checkbox.checked = false;
    });
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, checks[0] ?? okBtn);
  });
