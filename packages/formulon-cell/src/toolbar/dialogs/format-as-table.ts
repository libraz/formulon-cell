import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import {
  appendDialogActions,
  appendErrorRow,
  appendInputRow,
  createDialogShell,
  focusAndSelectInput,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface FormatAsTableDialogOptions {
  title: string;
  rangeLabel: string;
  headersLabel: string;
  initialRange: string;
  initialHasHeaders: boolean;
  okLabel: string;
  cancelLabel: string;
  validateRange: (value: string) => string | null;
  /** Table-name field, shown when a label is given. */
  nameLabel?: string;
  initialName?: string;
  validateName?: (value: string) => string | null;
  rangePickerLabel?: string;
  pickRange?: () => string;
  subscribeToRangeChanges?: (listener: () => void) => () => void;
}

export interface FormatAsTableDialogResult {
  range: string;
  hasHeaders: boolean;
  /** Present when the dialog showed a name field. */
  name?: string;
}

export const showFormatAsTableDialog = (
  opts: FormatAsTableDialogOptions,
): Promise<FormatAsTableDialogResult | null> =>
  new Promise<FormatAsTableDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });
    const rangeInput = appendInputRow(shell.body, opts.rangeLabel, { initial: opts.initialRange });
    if (opts.pickRange) {
      attachRangePickerButton(rangeInput, {
        label: opts.rangePickerLabel ?? opts.rangeLabel,
        getValue: opts.pickRange,
        subscribeToRangeChanges: opts.subscribeToRangeChanges,
        kind: 'table-range',
      });
    }

    const nameInput = opts.nameLabel
      ? appendInputRow(shell.body, opts.nameLabel, { initial: opts.initialName ?? '' })
      : null;

    const headersLabel = document.createElement('label');
    headersLabel.className = 'fc-fmtdlg__check fc-tb__dlg__check';
    const headersInput = document.createElement('input');
    headersInput.type = 'checkbox';
    headersInput.checked = opts.initialHasHeaders;
    const headersText = document.createElement('span');
    headersText.textContent = opts.headersLabel;
    headersLabel.append(headersInput, headersText);
    shell.body.appendChild(headersLabel);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });

    const lifecycle = installDialogLifecycle<FormatAsTableDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const range = rangeInput.value.trim();
      const err = opts.validateRange(range);
      if (err) {
        showInputError(errorRow, rangeInput, err);
        return;
      }
      if (!nameInput) {
        lifecycle.finish({ range, hasHeaders: headersInput.checked });
        return;
      }
      const name = nameInput.value.trim();
      const nameErr = opts.validateName?.(name) ?? null;
      if (nameErr) {
        showInputError(errorRow, nameInput, nameErr);
        return;
      }
      lifecycle.finish({ range, hasHeaders: headersInput.checked, name });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(rangeInput));
  });
