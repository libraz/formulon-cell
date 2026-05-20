import {
  appendDialogActions,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';
import { attachRangePickerButton } from '../../interact/range-picker-control.js';

export interface AdvancedFilterDialogOptions {
  title: string;
  listRangeLabel: string;
  criteriaRangeLabel: string;
  copyToLabel: string;
  uniqueOnlyLabel: string;
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
  label.className = `fc-advfilter__row app__dlg__label ${className}`;
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
    if (opts.rangePickerLabel && pickCopyTo) {
      attachRangePickerButton(copyTo, {
        label: opts.rangePickerLabel,
        getValue: pickCopyTo,
        subscribeToRangeChanges: opts.subscribeToRangeChanges,
        kind: 'advanced-filter-copy-to',
      });
    }

    const uniqueRow = document.createElement('label');
    uniqueRow.className = 'fc-advfilter__option app__dlg__label';
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
      const copyError = copyTo.value.trim() ? opts.validateAddress(copyTo.value) : null;
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
        copyTo: copyTo.value.trim(),
        uniqueOnly: uniqueOnly.checked,
      });
    };

    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, listRange);
  });
