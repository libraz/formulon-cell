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
  okLabel?: string;
  cancelLabel?: string;
}

/** Excel-style modal radio-choice dialog. Used for commands like Insert/Delete
 *  Cells where desktop Excel presents a compact option list, not text input. */
export const showChoiceDialog = <T extends string>(
  opts: ChoiceDialogOptions<T>,
): Promise<T | null> =>
  new Promise<T | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });

    if (opts.label) {
      const label = document.createElement('p');
      label.className = 'app__dlg__message';
      label.textContent = opts.label;
      shell.body.appendChild(label);
    }

    const group = document.createElement('div');
    group.className = 'fc-fmtdlg__choice-grid app__dlg__choices';
    group.setAttribute('role', 'radiogroup');
    group.setAttribute('aria-label', opts.label ?? opts.title);
    const name = `app-choice-${Math.random().toString(36).slice(2)}`;
    for (const [index, option] of opts.options.entries()) {
      const wrap = document.createElement('label');
      wrap.className = 'fc-fmtdlg__radio';
      const input = document.createElement('input');
      input.type = 'radio';
      input.name = name;
      input.value = option.value;
      input.checked = option.value === opts.initial || (!opts.initial && index === 0);
      const text = document.createElement('span');
      text.textContent = option.label;
      wrap.append(input, text);
      group.appendChild(wrap);
    }
    shell.body.appendChild(group);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel ?? 'Cancel',
      okLabel: opts.okLabel ?? 'OK',
    });

    const selected = (): T | null =>
      (group.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.value as
        | T
        | undefined) ?? null;
    const lifecycle = installDialogLifecycle<T | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(selected()),
    });
    okBtn.addEventListener('click', () => lifecycle.finish(selected()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () =>
      group.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.focus(),
    );
  });
