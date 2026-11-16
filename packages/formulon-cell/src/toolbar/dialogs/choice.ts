import { appendRadioGroup } from './form-controls.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface ChoiceDialogOption<T extends string = string> {
  value: T;
  label: string;
}

export interface ChoiceDialogOptions<T extends string = string> {
  title: string;
  label?: string;
  options: readonly ChoiceDialogOption<T>[];
  initial?: T;
  okLabel: string;
  cancelLabel: string;
}

/** Excel-style modal radio-choice dialog. Used for commands like Insert/Delete
 *  Cells where desktop Excel presents a compact option list, not text input. */
export const showChoiceDialog = <T extends string>(
  opts: ChoiceDialogOptions<T>,
): Promise<T | null> =>
  new Promise<T | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });

    const { group, read } = appendRadioGroup(shell.body, {
      ariaLabel: opts.label ?? opts.title,
      caption: opts.label || undefined,
      options: opts.options,
      isChecked: (value, index) => value === opts.initial || (!opts.initial && index === 0),
    });

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });

    const selected = (): T | null => read() ?? null;
    const lifecycle = installDialogLifecycle<T | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(selected()),
      buttons: { okBtn, cancelBtn },
    });

    mountDialog(shell, () =>
      group.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.focus(),
    );
  });
