import type { PhoneticSettings } from '../../commands/phonetic.js';
import type { PhoneticAlignment } from '../../engine/phonetic-sync.js';
import { appendRadioGroup } from './form-controls.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface PhoneticSettingsDialogOptions {
  title: string;
  typeLabel: string;
  alignmentLabel: string;
  typeOptions: readonly { value: PhoneticSettings['form']; label: string }[];
  alignmentOptions: readonly { value: PhoneticAlignment; label: string }[];
  initial: PhoneticSettings;
  okLabel: string;
  cancelLabel: string;
}

const appendPhoneticRadioGroup = <T extends string>(
  body: HTMLElement,
  label: string,
  options: readonly { value: T; label: string }[],
  initial: T,
): (() => T) => {
  const { read } = appendRadioGroup(body, {
    ariaLabel: label,
    caption: label,
    options,
    isChecked: (value) => value === initial,
  });
  return () => read() ?? initial;
};

/** Phonetic guide settings: kana form and alignment. Returns the choice, or
 *  `null` when the user cancels. */
export const showPhoneticSettingsDialog = (
  opts: PhoneticSettingsDialogOptions,
): Promise<PhoneticSettings | null> =>
  new Promise<PhoneticSettings | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });
    const readForm = appendPhoneticRadioGroup(
      shell.body,
      opts.typeLabel,
      opts.typeOptions,
      opts.initial.form,
    );
    const readAlignment = appendPhoneticRadioGroup(
      shell.body,
      opts.alignmentLabel,
      opts.alignmentOptions,
      opts.initial.alignment,
    );
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });
    const read = (): PhoneticSettings => ({ form: readForm(), alignment: readAlignment() });
    const lifecycle = installDialogLifecycle<PhoneticSettings | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(read()),
      buttons: { okBtn, cancelBtn },
    });
    mountDialog(shell, () =>
      shell.body.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.focus(),
    );
  });
