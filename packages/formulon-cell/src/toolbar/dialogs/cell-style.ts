import type { CellStyleIncludeOptions } from '../../commands/cell-styles.js';
import { appendCheckboxRow } from './form-controls.js';
import {
  appendDialogActions,
  appendErrorRow,
  appendInputRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface CellStyleDialogResult {
  name: string;
  include: Required<CellStyleIncludeOptions>;
}

interface CellStyleDialogStrings {
  ribbonMenu: { [key: string]: string | undefined };
  hyperlinkDialog: { cancel: string; ok: string };
  namedRangeDialog: { errorEmptyName: string };
}

export interface CellStyleDialogOptions {
  title: string;
  strings: CellStyleDialogStrings;
  initialName?: string;
  initialInclude?: CellStyleIncludeOptions;
}

const DEFAULT_INCLUDE: Required<CellStyleIncludeOptions> = {
  number: true,
  alignment: true,
  font: true,
  border: true,
  fill: true,
  protection: true,
};

const menuText = (
  strings: CellStyleDialogStrings['ribbonMenu'],
  key: string,
  fallback: string,
): string => strings[key] ?? fallback;

export const showCellStyleDialog = (
  opts: CellStyleDialogOptions,
): Promise<CellStyleDialogResult | null> =>
  new Promise<CellStyleDialogResult | null>((resolve) => {
    const t = opts.strings.ribbonMenu;
    const include = { ...DEFAULT_INCLUDE, ...opts.initialInclude };
    const shell = createDialogShell({ title: opts.title });
    const name = appendInputRow(shell.body, menuText(t, 'cellStyleName', 'Style name'), {
      initial: opts.initialName ?? menuText(t, 'cellStyleNormal', 'Normal'),
    });
    name.dataset.dialogField = 'name';

    const legend = document.createElement('div');
    legend.className = 'app__dlg__label';
    legend.textContent = menuText(t, 'cellStyleIncludes', 'Style includes');
    shell.body.appendChild(legend);

    const number = appendCheckboxRow(
      shell.body,
      menuText(t, 'cellStyleIncludeNumber', 'Number'),
      include.number,
      'number',
    );
    const alignment = appendCheckboxRow(
      shell.body,
      menuText(t, 'cellStyleIncludeAlignment', 'Alignment'),
      include.alignment,
      'alignment',
    );
    const font = appendCheckboxRow(
      shell.body,
      menuText(t, 'cellStyleIncludeFont', 'Font'),
      include.font,
      'font',
    );
    const border = appendCheckboxRow(
      shell.body,
      menuText(t, 'cellStyleIncludeBorder', 'Border'),
      include.border,
      'border',
    );
    const fill = appendCheckboxRow(
      shell.body,
      menuText(t, 'cellStyleIncludeFill', 'Fill'),
      include.fill,
      'fill',
    );
    const protection = appendCheckboxRow(
      shell.body,
      menuText(t, 'cellStyleIncludeProtection', 'Protection'),
      include.protection,
      'protection',
    );
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.strings.hyperlinkDialog.cancel,
      okLabel: opts.strings.hyperlinkDialog.ok,
    });

    const lifecycle = installDialogLifecycle<CellStyleDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): CellStyleDialogResult | null => {
      const label = name.value.trim();
      if (!label) {
        showInputError(errorRow, name, opts.strings.namedRangeDialog.errorEmptyName);
        return null;
      }
      const result: CellStyleDialogResult = {
        name: label,
        include: {
          number: number.checked,
          alignment: alignment.checked,
          font: font.checked,
          border: border.checked,
          fill: fill.checked,
          protection: protection.checked,
        },
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
