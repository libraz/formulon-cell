import type { CellStyleIncludeOptions } from '../../commands/cell-styles.js';
import { appendCheckboxRow } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogNameField,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface CellStyleDialogResult {
  name: string;
  include: Required<CellStyleIncludeOptions>;
}

interface CellStyleDialogStrings {
  ribbonMenu: {
    cellStyleName: string;
    cellStyleNormal: string;
    cellStyleIncludes: string;
    cellStyleIncludeNumber: string;
    cellStyleIncludeAlignment: string;
    cellStyleIncludeFont: string;
    cellStyleIncludeBorder: string;
    cellStyleIncludeFill: string;
    cellStyleIncludeProtection: string;
  };
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

export const showCellStyleDialog = (
  opts: CellStyleDialogOptions,
): Promise<CellStyleDialogResult | null> =>
  new Promise<CellStyleDialogResult | null>((resolve) => {
    const t = opts.strings.ribbonMenu;
    const include = { ...DEFAULT_INCLUDE, ...opts.initialInclude };
    const shell = createDialogShell({ title: opts.title });
    const name = appendDialogNameField(
      shell.body,
      t.cellStyleName,
      opts.initialName ?? t.cellStyleNormal,
    );

    const legend = document.createElement('div');
    legend.className = 'fc-tb__dlg__label';
    legend.textContent = t.cellStyleIncludes;
    shell.body.appendChild(legend);

    const number = appendCheckboxRow(
      shell.body,
      t.cellStyleIncludeNumber,
      include.number,
      'number',
    );
    const alignment = appendCheckboxRow(
      shell.body,
      t.cellStyleIncludeAlignment,
      include.alignment,
      'alignment',
    );
    const font = appendCheckboxRow(shell.body, t.cellStyleIncludeFont, include.font, 'font');
    const border = appendCheckboxRow(
      shell.body,
      t.cellStyleIncludeBorder,
      include.border,
      'border',
    );
    const fill = appendCheckboxRow(shell.body, t.cellStyleIncludeFill, include.fill, 'fill');
    const protection = appendCheckboxRow(
      shell.body,
      t.cellStyleIncludeProtection,
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
      const label = name.valueOrError(errorRow, opts.strings.namedRangeDialog.errorEmptyName);
      if (label === null) {
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

    mountDialog(shell, name.focus);
  });
