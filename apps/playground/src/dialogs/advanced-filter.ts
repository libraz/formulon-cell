import {
  appendDialogActions,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface AdvancedFilterDialogOptions {
  title: string;
  listRangeLabel: string;
  criteriaRangeLabel: string;
  copyToLabel: string;
  uniqueOnlyLabel: string;
  initialListRange: string;
  okLabel?: string;
  cancelLabel?: string;
  validateListRange: (value: string) => string | null;
  validateCriteriaRange: (value: string) => string | null;
  validateCopyTo: (value: string) => string | null;
}

export interface AdvancedFilterDialogResult {
  listRange: string;
  criteriaRange: string;
  copyTo: string;
  uniqueOnly: boolean;
}

export const showAdvancedFilterDialog = (
  opts: AdvancedFilterDialogOptions,
): Promise<AdvancedFilterDialogResult | null> =>
  new Promise<AdvancedFilterDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });

    const makeInput = (labelText: string, value = ''): HTMLInputElement => {
      const row = document.createElement('div');
      row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
      const label = document.createElement('label');
      label.className = 'app__dlg__label';
      label.textContent = labelText;
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'app__dlg__input';
      input.value = value;
      label.appendChild(input);
      row.appendChild(label);
      shell.body.appendChild(row);
      return input;
    };

    const listInput = makeInput(opts.listRangeLabel, opts.initialListRange);
    const criteriaInput = makeInput(opts.criteriaRangeLabel);
    const copyInput = makeInput(opts.copyToLabel);

    const uniqueLabel = document.createElement('label');
    uniqueLabel.className = 'fc-fmtdlg__check app__dlg__check';
    const uniqueInput = document.createElement('input');
    uniqueInput.type = 'checkbox';
    const uniqueText = document.createElement('span');
    uniqueText.textContent = opts.uniqueOnlyLabel;
    uniqueLabel.append(uniqueInput, uniqueText);
    shell.body.appendChild(uniqueLabel);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel ?? 'Cancel',
      okLabel: opts.okLabel ?? 'OK',
    });

    const lifecycle = installDialogLifecycle<AdvancedFilterDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const listRange = listInput.value.trim();
      const criteriaRange = criteriaInput.value.trim();
      const copyTo = copyInput.value.trim();
      const listError = opts.validateListRange(listRange);
      if (listError) {
        showInputError(errorRow, listInput, listError);
        return;
      }
      const criteriaError = opts.validateCriteriaRange(criteriaRange);
      if (criteriaError) {
        showInputError(errorRow, criteriaInput, criteriaError);
        return;
      }
      const copyError = opts.validateCopyTo(copyTo);
      if (copyError) {
        showInputError(errorRow, copyInput, copyError);
        return;
      }
      lifecycle.finish({ listRange, criteriaRange, copyTo, uniqueOnly: uniqueInput.checked });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => {
      listInput.focus();
      listInput.select();
    });
  });
