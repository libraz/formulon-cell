import type { DefinedNameEntry } from '../../commands/named-ranges.js';
import { appendNamedRadioRow } from './form-controls.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface DefinedNamePickerDialogOptions {
  title: string;
  names: readonly DefinedNameEntry[];
  okLabel: string;
  cancelLabel: string;
}

export const showDefinedNamePickerDialog = (
  opts: DefinedNamePickerDialogOptions,
): Promise<string | null> =>
  new Promise<string | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });
    const list = document.createElement('div');
    list.className = 'fc-tb__dlg__list';
    list.setAttribute('role', 'radiogroup');
    list.setAttribute('aria-label', opts.title);

    let selected = opts.names[0]?.name ?? '';
    const radios: HTMLInputElement[] = [];
    for (const [index, entry] of opts.names.entries()) {
      radios.push(
        appendNamedRadioRow(list, {
          name: 'fc-defined-name-picker',
          value: entry.name,
          title: entry.name,
          note: entry.formula,
          checked: index === 0,
          onSelect: (value) => {
            selected = value;
          },
        }),
      );
    }
    shell.body.appendChild(list);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });
    const lifecycle = installDialogLifecycle<string | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      lifecycle.finish(selected || null);
    };

    mountDialog(shell, () => {
      (radios[0] ?? okBtn).focus();
    });
  });
