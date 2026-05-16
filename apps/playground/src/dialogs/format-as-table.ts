import {
  appendDialogActions,
  appendErrorRow,
  appendInputRow,
  createDialogShell,
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
  okLabel?: string;
  cancelLabel?: string;
  validateRange: (value: string) => string | null;
}

export interface FormatAsTableDialogResult {
  range: string;
  hasHeaders: boolean;
}

export const showFormatAsTableDialog = (
  opts: FormatAsTableDialogOptions,
): Promise<FormatAsTableDialogResult | null> =>
  new Promise<FormatAsTableDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });
    const rangeInput = appendInputRow(shell.body, opts.rangeLabel, { initial: opts.initialRange });

    const headersLabel = document.createElement('label');
    headersLabel.className = 'fc-fmtdlg__check app__dlg__check';
    const headersInput = document.createElement('input');
    headersInput.type = 'checkbox';
    headersInput.checked = opts.initialHasHeaders;
    const headersText = document.createElement('span');
    headersText.textContent = opts.headersLabel;
    headersLabel.append(headersInput, headersText);
    shell.body.appendChild(headersLabel);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel ?? 'Cancel',
      okLabel: opts.okLabel ?? 'OK',
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
      lifecycle.finish({ range, hasHeaders: headersInput.checked });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => {
      rangeInput.focus();
      rangeInput.select();
    });
  });
