import { appendSelectRow } from './form-controls.js';
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
    const select = appendSelectRow(
      shell.body,
      opts.label,
      opts.options,
      opts.initial ?? opts.options[0]?.value ?? '',
      undefined,
      'fc-tb__dlg__select',
    );
    select.dataset.scriptCommandSelect = 'true';

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });
    const lifecycle = installDialogLifecycle<T | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      const value = select.value as T;
      lifecycle.finish(value);
    };

    mountDialog(shell, select);
  });
