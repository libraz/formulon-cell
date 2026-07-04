import type { DefinedNameEntry } from '../../commands/named-ranges.js';
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
      const label = document.createElement('label');
      label.className = 'fc-fmtdlg__row fc-fmtdlg__row--block fc-tb__dlg__check';

      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'fc-defined-name-picker';
      radio.value = entry.name;
      radio.checked = index === 0;
      radio.addEventListener('change', () => {
        if (radio.checked) selected = radio.value;
      });
      radios.push(radio);

      const body = document.createElement('span');
      const name = document.createElement('strong');
      name.textContent = entry.name;
      const formula = document.createElement('span');
      formula.className = 'fc-tb__dlg__note';
      formula.textContent = entry.formula;
      body.append(name, formula);

      label.append(radio, body);
      list.appendChild(label);
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
    });
    const onOk = (): void => {
      lifecycle.finish(selected || null);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => {
      (radios[0] ?? okBtn).focus();
    });
  });
