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
  /** The spreadsheet host; anchors the dialog and picks its platform layout. */
  host?: HTMLElement;
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
    const shell = createDialogShell({ host: opts.host, title: opts.title, bodyVariant: 'app' });
    // The Mac app puts a tri-state Select All row at the top of the list in
    // place of the Select All / Unselect All buttons.
    const mac =
      shell.overlay.closest<HTMLElement>('[data-fc-platform]')?.dataset.fcPlatform === 'mac';
    shell.panel.classList.add('fc-dedupedlg');

    const headerRow = document.createElement('label');
    headerRow.className = 'fc-fmtdlg__row fc-tb__dlg__label fc-dedupedlg__header';
    const hasHeader = document.createElement('input');
    hasHeader.type = 'checkbox';
    hasHeader.checked = opts.initialHasHeader;
    headerRow.append(hasHeader, document.createTextNode(` ${opts.headerLabel}`));
    shell.body.appendChild(headerRow);

    const setAll = (checked: boolean): void => {
      for (const checkbox of checks) checkbox.checked = checked;
      if (checked) clearDialogError(errorRow);
      syncSelectAll();
    };
    if (!mac) {
      const actionRow = document.createElement('div');
      actionRow.className = 'fc-dedupedlg__actions';
      const selectAllBtn = appendDialogButton(actionRow, { label: opts.selectAllLabel });
      const unselectAllBtn = appendDialogButton(actionRow, { label: opts.unselectAllLabel });
      shell.body.appendChild(actionRow);
      selectAllBtn.addEventListener('click', () => setAll(true));
      unselectAllBtn.addEventListener('click', () => setAll(false));
    }

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
    const labels: { text: HTMLSpanElement; item: SortDialogColumn }[] = [];
    const initialColumns = new Set(opts.initialColumns);
    const row = (
      text: string,
    ): { label: HTMLLabelElement; checkbox: HTMLInputElement; span: HTMLSpanElement } => {
      const label = document.createElement('label');
      label.className = 'fc-dedupedlg__column';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      const span = document.createElement('span');
      span.textContent = text;
      label.append(checkbox, span);
      list.appendChild(label);
      return { label, checkbox, span };
    };
    const selectAll = mac ? row(opts.selectAllLabel) : null;
    selectAll?.label.classList.add('fc-dedupedlg__select-all');
    for (const item of opts.columns) {
      const { checkbox, span } = row(item.label);
      checkbox.value = item.value;
      checkbox.checked = initialColumns.has(item.value);
      checks.push(checkbox);
      labels.push({ text: span, item });
    }
    shell.body.appendChild(fieldset);
    // Header text names the columns while the first row counts as headers.
    const relabel = (): void => {
      for (const { text, item } of labels) {
        text.textContent = hasHeader.checked && item.heading ? item.heading : item.label;
      }
    };
    function syncSelectAll(): void {
      if (!selectAll) return;
      const on = checks.filter((checkbox) => checkbox.checked).length;
      selectAll.checkbox.checked = on === checks.length && on > 0;
      selectAll.checkbox.indeterminate = on > 0 && on < checks.length;
    }
    relabel();
    syncSelectAll();
    hasHeader.addEventListener('change', relabel);
    selectAll?.checkbox.addEventListener('change', () => setAll(selectAll.checkbox.checked));
    for (const checkbox of checks) checkbox.addEventListener('change', syncSelectAll);

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
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, checks[0] ?? okBtn);
  });
