import type { PhoneticSettings } from '../../commands/phonetic.js';
import type { PhoneticAlignment } from '../../engine/phonetic-sync.js';
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

const appendRadioGroup = <T extends string>(
  body: HTMLElement,
  label: string,
  options: readonly { value: T; label: string }[],
  initial: T,
): (() => T) => {
  const group = document.createElement('div');
  group.className = 'fc-fmtdlg__choice-grid fc-tb__dlg__choices';
  group.setAttribute('role', 'radiogroup');
  group.setAttribute('aria-label', label);
  const caption = document.createElement('p');
  caption.className = 'fc-tb__dlg__message';
  caption.textContent = label;
  const name = `phonetic-${Math.random().toString(36).slice(2)}`;
  for (const option of options) {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = name;
    input.value = option.value;
    input.checked = option.value === initial;
    const text = document.createElement('span');
    text.textContent = option.label;
    wrap.append(input, text);
    group.appendChild(wrap);
  }
  body.append(caption, group);
  return () =>
    (group.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.value as
      | T
      | undefined) ?? initial;
};

/** Phonetic guide settings: kana form and alignment. Returns the choice, or
 *  `null` when the user cancels. */
export const showPhoneticSettingsDialog = (
  opts: PhoneticSettingsDialogOptions,
): Promise<PhoneticSettings | null> =>
  new Promise<PhoneticSettings | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });
    const readForm = appendRadioGroup(
      shell.body,
      opts.typeLabel,
      opts.typeOptions,
      opts.initial.form,
    );
    const readAlignment = appendRadioGroup(
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
    });
    okBtn.addEventListener('click', () => lifecycle.finish(read()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));
    mountDialog(shell, () =>
      shell.body.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.focus(),
    );
  });
