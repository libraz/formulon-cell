import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface SortDialogColumn {
  value: string;
  label: string;
}

export interface SortDialogOptions {
  title: string;
  columnLabel: string;
  thenByLabel?: string;
  noThenByLabel?: string;
  orderLabel: string;
  headerLabel: string;
  ascendingLabel: string;
  descendingLabel: string;
  columns: readonly SortDialogColumn[];
  initialColumn: string;
  initialDirection: 'asc' | 'desc';
  initialHasHeader: boolean;
  okLabel?: string;
  cancelLabel?: string;
}

export interface SortDialogResult {
  column: string;
  direction: 'asc' | 'desc';
  levels: Array<{ column: string; direction: 'asc' | 'desc' }>;
  hasHeader: boolean;
}

const buildColumnSelect = (
  columns: readonly SortDialogColumn[],
  initial: string,
): HTMLSelectElement => {
  const select = document.createElement('select');
  select.className = 'app__dlg__input';
  for (const item of columns) {
    const option = document.createElement('option');
    option.value = item.value;
    option.textContent = item.label;
    select.appendChild(option);
  }
  select.value = initial;
  return select;
};

const buildDirectionSelect = (
  asc: string,
  desc: string,
  initial: 'asc' | 'desc',
): HTMLSelectElement => {
  const select = document.createElement('select');
  select.className = 'app__dlg__input';
  const ascOpt = document.createElement('option');
  ascOpt.value = 'asc';
  ascOpt.textContent = asc;
  select.appendChild(ascOpt);
  const descOpt = document.createElement('option');
  descOpt.value = 'desc';
  descOpt.textContent = desc;
  select.appendChild(descOpt);
  select.value = initial;
  return select;
};

const appendSelectRow = (body: HTMLElement, labelText: string, select: HTMLSelectElement): void => {
  const row = document.createElement('label');
  row.className = 'fc-fmtdlg__row app__dlg__label';
  row.textContent = labelText;
  row.appendChild(select);
  body.appendChild(row);
};

export const showSortDialog = (opts: SortDialogOptions): Promise<SortDialogResult | null> =>
  new Promise<SortDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });

    const column = buildColumnSelect(opts.columns, opts.initialColumn);
    appendSelectRow(shell.body, opts.columnLabel, column);

    const thenColumn = document.createElement('select');
    thenColumn.className = 'app__dlg__input';
    const noThen = document.createElement('option');
    noThen.value = '';
    noThen.textContent = opts.noThenByLabel ?? '(none)';
    thenColumn.appendChild(noThen);
    for (const item of opts.columns) {
      const option = document.createElement('option');
      option.value = item.value;
      option.textContent = item.label;
      thenColumn.appendChild(option);
    }
    appendSelectRow(shell.body, opts.thenByLabel ?? 'Then by', thenColumn);

    const order = buildDirectionSelect(
      opts.ascendingLabel,
      opts.descendingLabel,
      opts.initialDirection,
    );
    appendSelectRow(shell.body, opts.orderLabel, order);

    const thenOrder = buildDirectionSelect(opts.ascendingLabel, opts.descendingLabel, 'asc');
    thenOrder.disabled = true;
    appendSelectRow(shell.body, `${opts.thenByLabel ?? 'Then by'} ${opts.orderLabel}`, thenOrder);
    thenColumn.addEventListener('change', () => {
      thenOrder.disabled = !thenColumn.value;
    });

    const headerRow = document.createElement('label');
    headerRow.className = 'fc-fmtdlg__row app__dlg__label';
    const hasHeader = document.createElement('input');
    hasHeader.type = 'checkbox';
    hasHeader.checked = opts.initialHasHeader;
    headerRow.append(hasHeader, document.createTextNode(` ${opts.headerLabel}`));
    shell.body.appendChild(headerRow);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel ?? 'Cancel',
      okLabel: opts.okLabel ?? 'OK',
    });

    const buildResult = (): SortDialogResult => ({
      column: column.value,
      direction: order.value === 'desc' ? 'desc' : 'asc',
      levels: [
        { column: column.value, direction: order.value === 'desc' ? 'desc' : 'asc' },
        ...(thenColumn.value
          ? [
              {
                column: thenColumn.value,
                direction: thenOrder.value === 'desc' ? ('desc' as const) : ('asc' as const),
              },
            ]
          : []),
      ],
      hasHeader: hasHeader.checked,
    });

    const lifecycle = installDialogLifecycle<SortDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(buildResult()),
    });
    okBtn.addEventListener('click', () => lifecycle.finish(buildResult()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, column);
  });
