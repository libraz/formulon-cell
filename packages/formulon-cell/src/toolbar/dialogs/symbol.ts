import type { ToolbarMenuText } from '../../index.js';

import { toolbarSymbolGroups } from '../ribbon/symbols.js';
import {
  appendDialogActions,
  appendErrorRow,
  appendInputRow,
  createDialogChoiceButton,
  createDialogShell,
  focusAndSelectInput,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface SymbolDialogOptions {
  text: ToolbarMenuText;
  okLabel: string;
  cancelLabel: string;
}

export const showSymbolDialog = (opts: SymbolDialogOptions): Promise<string | null> =>
  new Promise<string | null>((resolve) => {
    const shell = createDialogShell({ title: opts.text.symbolMore, bodyVariant: 'app' });
    const input = appendInputRow(shell.body, opts.text.symbolPrompt, {
      initial: '',
      placeholder: opts.text.symbol,
    });
    const errorRow = appendErrorRow(shell.body);

    const picker = document.createElement('div');
    picker.className = 'app__symbol-dialog';
    picker.setAttribute('role', 'group');
    picker.setAttribute('aria-label', opts.text.symbol);
    for (const group of toolbarSymbolGroups(opts.text)) {
      const heading = document.createElement('div');
      heading.className = 'app__menu-heading';
      heading.textContent = group.label;
      picker.appendChild(heading);

      const row = document.createElement('div');
      row.className = 'app__symbol-dialog__row';
      for (const symbol of group.symbols) {
        const button = createDialogChoiceButton({ label: symbol, title: symbol });
        button.addEventListener('click', () => {
          input.value = symbol;
          focusAndSelectInput(input);
        });
        row.appendChild(button);
      }
      picker.appendChild(row);
    }
    shell.body.appendChild(picker);

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
        showInputError(errorRow, input, opts.text.symbolInvalid);
        return;
      }
      lifecycle.finish(value);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, input);
  });
