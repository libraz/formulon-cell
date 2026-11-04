import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import { projectDisabledState } from '../menu-a11y.js';
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
  /** In-place / copy action radios; Copy to stays disabled while in place. */
  actionLabels?: { inPlace: string; copy: string };
  initialListRange: string;
  initialCriteriaRange?: string;
  initialCopyTo?: string;
  okLabel: string;
  cancelLabel: string;
  rangePickerLabel?: string;
  pickRange?: () => string;
  pickAddress?: () => string;
  subscribeToRangeChanges?: (listener: () => void) => () => void;
  validateRange: (value: string) => string | null;
  validateAddress: (value: string) => string | null;
}

export interface AdvancedFilterDialogResult {
  listRange: string;
  criteriaRange: string;
  copyTo: string;
  uniqueOnly: boolean;
}

const appendTextRow = (
  parent: HTMLElement,
  labelText: string,
  initial: string,
  className: string,
): HTMLInputElement => {
  const label = document.createElement('label');
  label.className = `fc-advfilter__row fc-tb__dlg__label ${className}`;
  const span = document.createElement('span');
  span.textContent = labelText;
  const input = document.createElement('input');
  input.className = 'fc-namedlg__input';
  input.type = 'text';
  input.value = initial;
  label.append(span, input);
  parent.appendChild(label);
  return input;
};

export const showAdvancedFilterDialog = (
  opts: AdvancedFilterDialogOptions,
): Promise<AdvancedFilterDialogResult | null> =>
  new Promise<AdvancedFilterDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });

    const actionInputs = opts.actionLabels
      ? [opts.actionLabels.inPlace, opts.actionLabels.copy].map((text, index) => {
          const label = document.createElement('label');
          label.className = 'fc-advfilter__option fc-tb__dlg__label';
          const input = document.createElement('input');
          input.type = 'radio';
          input.name = 'fc-advfilter-action';
          input.value = index === 0 ? 'in-place' : 'copy';
          input.checked = index === 0;
          label.append(input, document.createTextNode(text));
          shell.body.appendChild(label);
          return input;
        })
      : null;

    const rangeGroup = document.createElement('div');
    rangeGroup.className = 'fc-advfilter__ranges';
    shell.body.appendChild(rangeGroup);

    const listRange = appendTextRow(
      rangeGroup,
      opts.listRangeLabel,
      opts.initialListRange,
      'fc-advfilter__row--list',
    );
    if (opts.rangePickerLabel && opts.pickRange) {
      attachRangePickerButton(listRange, {
        label: opts.rangePickerLabel,
        getValue: opts.pickRange,
        subscribeToRangeChanges: opts.subscribeToRangeChanges,
        kind: 'advanced-filter-list-range',
      });
    }
    const criteriaRange = appendTextRow(
      rangeGroup,
      opts.criteriaRangeLabel,
      opts.initialCriteriaRange ?? '',
      'fc-advfilter__row--criteria',
    );
    if (opts.rangePickerLabel && opts.pickRange) {
      attachRangePickerButton(criteriaRange, {
        label: opts.rangePickerLabel,
        getValue: opts.pickRange,
        subscribeToRangeChanges: opts.subscribeToRangeChanges,
        kind: 'advanced-filter-criteria-range',
      });
    }
    const copyTo = appendTextRow(
      rangeGroup,
      opts.copyToLabel,
      opts.initialCopyTo ?? '',
      'fc-advfilter__row--copy-to',
    );
    const pickCopyTo = opts.pickAddress ?? opts.pickRange;
    let copyToPicker: HTMLButtonElement | null = null;
    if (opts.rangePickerLabel && pickCopyTo) {
      copyToPicker = attachRangePickerButton(copyTo, {
        label: opts.rangePickerLabel,
        getValue: pickCopyTo,
        subscribeToRangeChanges: opts.subscribeToRangeChanges,
        kind: 'advanced-filter-copy-to',
      });
    }

    const copying = (): boolean => !actionInputs || actionInputs[1]?.checked === true;
    const syncCopyTo = (): void => {
      const disabled = !copying();
      projectDisabledState(copyTo, disabled, null);
      if (copyToPicker) projectDisabledState(copyToPicker, disabled, null);
    };
    for (const input of actionInputs ?? []) input.addEventListener('change', syncCopyTo);
    syncCopyTo();

    const uniqueRow = document.createElement('label');
    uniqueRow.className = 'fc-advfilter__option fc-tb__dlg__label';
    const uniqueOnly = document.createElement('input');
    uniqueOnly.type = 'checkbox';
    uniqueRow.append(uniqueOnly, document.createTextNode(` ${opts.uniqueOnlyLabel}`));
    shell.body.appendChild(uniqueRow);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });

    const lifecycle = installDialogLifecycle<AdvancedFilterDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });

    const onOk = (): void => {
      const listError = opts.validateRange(listRange.value);
      const criteriaError = opts.validateRange(criteriaRange.value);
      const copyValue = copying() ? copyTo.value.trim() : '';
      const copyError = copyValue ? opts.validateAddress(copyValue) : null;
      if (listError) {
        showInputError(errorRow, listRange, listError);
        return;
      }
      if (criteriaError) {
        showInputError(errorRow, criteriaRange, criteriaError);
        return;
      }
      if (copyError) {
        showInputError(errorRow, copyTo, copyError);
        return;
      }
      lifecycle.finish({
        listRange: listRange.value.trim(),
        criteriaRange: criteriaRange.value.trim(),
        copyTo: copyValue,
        uniqueOnly: uniqueOnly.checked,
      });
    };

    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, listRange);
  });
