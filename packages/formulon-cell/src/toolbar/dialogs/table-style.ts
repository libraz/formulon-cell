import {
  type CustomTableStyle,
  DEFAULT_TABLE_COLOR,
  TABLE_STYLE_COLORS,
  type TableStyle,
  tableVariantFromOptions,
  tableVariantOptions,
} from '../../commands/format-as-table.js';
import { appendCheckboxRow, appendSelectRow } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogNameField,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface TableStyleDialogResult {
  name: string;
  style: TableStyle;
  color: string;
  variant: CustomTableStyle['variant'];
}

interface TableStyleDialogStrings {
  ribbonMenu: {
    tableStyleName: string;
    tableStyleMedium: string;
    tableStyleType: string;
    tableStyleLight: string;
    tableStyleDark: string;
    tableStyleColor: string;
    tableStyleBandedRows: string;
    tableStyleFirstColumn: string;
  };
  hyperlinkDialog: { cancel: string; ok: string };
  namedRangeDialog: { errorEmptyName: string };
}

export interface TableStyleDialogOptions {
  title: string;
  strings: TableStyleDialogStrings;
  initial?: Partial<TableStyleDialogResult>;
}

export const showTableStyleDialog = (
  opts: TableStyleDialogOptions,
): Promise<TableStyleDialogResult | null> =>
  new Promise<TableStyleDialogResult | null>((resolve) => {
    const { strings } = opts;
    const t = strings.ribbonMenu;
    const initialVariant = tableVariantOptions(opts.initial?.variant ?? 'banded');
    const shell = createDialogShell({ title: opts.title });
    const name = appendDialogNameField(
      shell.body,
      t.tableStyleName,
      opts.initial?.name ?? t.tableStyleMedium,
    );
    const style = appendSelectRow(
      shell.body,
      t.tableStyleType,
      [
        { value: 'light', label: t.tableStyleLight },
        { value: 'medium', label: t.tableStyleMedium },
        { value: 'dark', label: t.tableStyleDark },
      ],
      opts.initial?.style ?? 'medium',
      'style',
    );
    const color = appendSelectRow(
      shell.body,
      t.tableStyleColor,
      TABLE_STYLE_COLORS.map((value) => ({ value, label: value })),
      opts.initial?.color ?? DEFAULT_TABLE_COLOR,
      'color',
    );
    const bandedRows = appendCheckboxRow(
      shell.body,
      t.tableStyleBandedRows,
      initialVariant.banded,
      'bandedRows',
    );
    const firstColumn = appendCheckboxRow(
      shell.body,
      t.tableStyleFirstColumn,
      initialVariant.firstCol,
      'firstColumn',
    );
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.hyperlinkDialog.cancel,
      okLabel: strings.hyperlinkDialog.ok,
    });

    const lifecycle = installDialogLifecycle<TableStyleDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): TableStyleDialogResult | null => {
      const label = name.valueOrError(errorRow, strings.namedRangeDialog.errorEmptyName);
      if (label === null) {
        return null;
      }
      const result: TableStyleDialogResult = {
        name: label,
        style: style.value as TableStyle,
        color: color.value,
        variant: tableVariantFromOptions({
          banded: bandedRows.checked,
          firstCol: firstColumn.checked,
        }),
      };
      lifecycle.finish(result);
      return result;
    };
    okBtn.addEventListener('click', () => {
      onOk();
    });
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, name.focus);
  });
