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

export interface RenameSheetDialogOptions {
  title: string;
  label: string;
  initial: string;
  requiredMessage: string;
  okLabel: string;
  cancelLabel: string;
}

export const showRenameSheetDialog = (opts: RenameSheetDialogOptions): Promise<string | null> =>
  new Promise<string | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    const input = appendInputRow(shell.body, opts.label, { initial: opts.initial });
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });

    const lifecycle = installDialogLifecycle<string | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const value = input.value.trim();
      if (!value) {
        showInputError(errorRow, input, opts.requiredMessage);
        return;
      }
      lifecycle.finish(value);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(input));
  });
