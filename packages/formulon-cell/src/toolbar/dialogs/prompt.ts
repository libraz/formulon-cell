import {
  appendDialogActions,
  appendDialogButton,
  appendErrorRow,
  appendInputRow,
  createDialogShell,
  focusAndSelectInput,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface PromptOptions {
  title: string;
  label: string;
  initial?: string;
  placeholder?: string;
  okLabel: string;
  cancelLabel: string;
  validate?: (value: string) => string | null;
}

export interface NumberPromptOptions {
  title: string;
  label: string;
  initial?: number;
  min?: number;
  max?: number;
  step?: number;
  okLabel: string;
  cancelLabel: string;
  invalidMessage: string;
}

export interface ConfirmOptions {
  title: string;
  message: string;
  okLabel: string;
  cancelLabel: string;
  destructive?: boolean;
}

export interface MessageOptions {
  title: string;
  message: string;
  okLabel: string;
}

/** Excel 365-styled modal prompt. Returns the entered value, or `null`
 *  when the user cancels. Replaces native `window.prompt`. */
export const showPrompt = (opts: PromptOptions): Promise<string | null> =>
  new Promise<string | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    const input = appendInputRow(shell.body, opts.label, {
      initial: opts.initial ?? '',
      placeholder: opts.placeholder,
    });
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
      const value = input.value;
      const err = opts.validate?.(value) ?? null;
      if (err) {
        showInputError(errorRow, input, err);
        return;
      }
      lifecycle.finish(value);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(input));
  });

/** Excel-styled numeric prompt. Returns a number, or `null` on cancel. */
export const showNumberPrompt = (opts: NumberPromptOptions): Promise<number | null> =>
  new Promise<number | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    const input = appendInputRow(shell.body, opts.label, {
      type: 'number',
      initial: Number.isFinite(opts.initial) ? String(opts.initial) : '',
      min: opts.min,
      max: opts.max,
      step: opts.step,
    });
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });

    const readValue = (): number | null => {
      const n = Number(input.value);
      if (!Number.isFinite(n)) return null;
      if (typeof opts.min === 'number' && n < opts.min) return null;
      if (typeof opts.max === 'number' && n > opts.max) return null;
      return n;
    };
    const lifecycle = installDialogLifecycle<number | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const n = readValue();
      if (n === null) {
        showInputError(errorRow, input, opts.invalidMessage);
        return;
      }
      lifecycle.finish(n);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(input));
  });

/** Excel 365-styled modal confirm. Returns true on accept, false on
 *  cancel/dismiss. Replaces native `window.confirm`. */
export const showConfirm = (opts: ConfirmOptions): Promise<boolean> =>
  new Promise<boolean>((resolve) => {
    const shell = createDialogShell({
      title: opts.title,
      role: 'alertdialog',
      bodyVariant: 'app',
    });
    const msg = document.createElement('p');
    msg.className = 'app__dlg__message';
    msg.textContent = opts.message;
    shell.body.appendChild(msg);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
      destructive: opts.destructive,
    });
    const lifecycle = installDialogLifecycle<boolean>({
      shell,
      resolve,
      onCancel: () => false,
      onSubmit: () => lifecycle.finish(true),
    });
    okBtn.addEventListener('click', () => lifecycle.finish(true));
    cancelBtn.addEventListener('click', () => lifecycle.finish(false));

    mountDialog(shell, okBtn);
  });

/** Excel 365-styled message dialog. Use for one-button errors/info instead
 *  of native `window.alert`, keeping focus and stacking inside app chrome. */
export const showMessage = (opts: MessageOptions): Promise<void> =>
  new Promise<void>((resolve) => {
    const shell = createDialogShell({
      title: opts.title,
      role: 'alertdialog',
      bodyVariant: 'app',
    });
    const msg = document.createElement('p');
    msg.className = 'app__dlg__message';
    msg.textContent = opts.message;
    shell.body.appendChild(msg);

    const okBtn = appendDialogButton(shell.footer, {
      label: opts.okLabel,
      variant: 'primary',
    });

    const lifecycle = installDialogLifecycle<void>({
      shell,
      resolve: () => resolve(),
      onCancel: () => undefined,
      onSubmit: () => lifecycle.finish(undefined),
    });
    okBtn.addEventListener('click', () => lifecycle.finish(undefined));

    mountDialog(shell, okBtn);
  });
