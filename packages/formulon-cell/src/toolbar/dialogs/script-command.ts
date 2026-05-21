import { createDialogSelect } from './form-controls.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface ScriptCommandDialogOption<T extends string> {
  value: T;
  label: string;
}

export interface ScriptCommandDialogOptions<T extends string> {
  title: string;
  label: string;
  options: readonly ScriptCommandDialogOption<T>[];
  initial?: T;
  okLabel: string;
  cancelLabel: string;
}

export const showScriptCommandDialog = <T extends string>(
  opts: ScriptCommandDialogOptions<T>,
): Promise<T | null> =>
  new Promise<T | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    const row = document.createElement('div');
    row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
    const label = document.createElement('label');
    label.className = 'app__dlg__label';
    label.textContent = opts.label;
    const select = createDialogSelect(opts.options, opts.initial ?? opts.options[0]?.value ?? '', {
      className: 'app__dlg__select',
    });
    select.dataset.scriptCommandSelect = 'true';
    label.appendChild(select);
    row.appendChild(label);
    shell.body.appendChild(row);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });
    const lifecycle = installDialogLifecycle<T | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const value = select.value as T;
      lifecycle.finish(value);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, select);
  });
