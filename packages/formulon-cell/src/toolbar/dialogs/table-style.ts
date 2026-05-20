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
  appendErrorRow,
  appendInputRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface TableStyleDialogResult {
  name: string;
  style: TableStyle;
  color: string;
  variant: CustomTableStyle['variant'];
}

interface TableStyleDialogStrings {
  ribbonMenu: { [key: string]: string | undefined };
  hyperlinkDialog: { cancel: string; ok: string };
  namedRangeDialog: { errorEmptyName: string };
}

export interface TableStyleDialogOptions {
  title: string;
  strings: TableStyleDialogStrings;
  initial?: Partial<TableStyleDialogResult>;
}

const menuText = (
  strings: TableStyleDialogStrings['ribbonMenu'],
  key: string,
  fallback: string,
): string => strings[key] ?? fallback;

export const showTableStyleDialog = (
  opts: TableStyleDialogOptions,
): Promise<TableStyleDialogResult | null> =>
  new Promise<TableStyleDialogResult | null>((resolve) => {
    const { strings } = opts;
    const t = strings.ribbonMenu;
    const initialVariant = tableVariantOptions(opts.initial?.variant ?? 'banded');
    const shell = createDialogShell({ title: opts.title });
    const name = appendInputRow(shell.body, menuText(t, 'tableStyleName', 'Style name'), {
      initial: opts.initial?.name ?? menuText(t, 'tableStyleMedium', 'Medium'),
    });
    name.dataset.dialogField = 'name';
    const style = appendSelectRow(
      shell.body,
      menuText(t, 'tableStyleType', 'Style type'),
      [
        { value: 'light', label: menuText(t, 'tableStyleLight', 'Light') },
        { value: 'medium', label: menuText(t, 'tableStyleMedium', 'Medium') },
        { value: 'dark', label: menuText(t, 'tableStyleDark', 'Dark') },
      ],
      opts.initial?.style ?? 'medium',
      'style',
    );
    const color = appendSelectRow(
      shell.body,
      menuText(t, 'tableStyleColor', 'Accent color'),
      TABLE_STYLE_COLORS.map((value) => ({ value, label: value })),
      opts.initial?.color ?? DEFAULT_TABLE_COLOR,
      'color',
    );
    const bandedRows = appendCheckboxRow(
      shell.body,
      menuText(t, 'tableStyleBandedRows', 'Banded rows'),
      initialVariant.banded,
      'bandedRows',
    );
    const firstColumn = appendCheckboxRow(
      shell.body,
      menuText(t, 'tableStyleFirstColumn', 'First column emphasis'),
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
      const label = name.value.trim();
      if (!label) {
        showInputError(errorRow, name, strings.namedRangeDialog.errorEmptyName);
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

    mountDialog(shell, () => {
      name.focus();
      name.select();
    });
  });
