import { parseLocalizedNumericInput } from '../../commands/coerce-input.js';
import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import {
  applyConditionalStylePreview,
  type ConditionalFormatDialogStyle,
  type ConditionalFormatStyleStrings,
  conditionalStyleFromValue,
  conditionalStyleOptions,
  DEFAULT_CF_STYLE_ID,
  showConditionalFormatCustomStyleDialog,
} from './conditional-format-style.js';
import { createDialogSelect } from './form-controls.js';
import {
  appendDialogActions,
  appendErrorRow,
  appendInputRow,
  createDialogShell,
  focusAndSelectInput,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface ConditionalFormatRuleDialogStrings extends ConditionalFormatStyleStrings {
  formatWith: string;
  formatPreview: string;
  invalidNumber: string;
  invalidText: string;
}

export interface ConditionalFormatDialogResult {
  /** Parsed number per field; `0` where the field holds a reference. */
  readonly values: readonly number[];
  /** Per field, the `=` reference text the user typed or picked, or `null` for a number. */
  readonly references?: readonly (string | null)[];
  readonly style: ConditionalFormatDialogStyle;
}

export interface ConditionalFormatTextDialogResult {
  readonly text: string;
  readonly style: ConditionalFormatDialogStyle;
}

export interface ConditionalFormatNumberDialogOptions {
  title: string;
  label: string;
  initial?: number;
  min?: number;
  max?: number;
  step?: number;
  secondLabel?: string;
  secondInitial?: number;
  initialStyle?: string;
  /** BCP 47 locale used by the decimal input boundary and initial value. */
  locale?: string;
  /** When present the fields also take a cell reference (`=B2`) and pick it from the grid. */
  reference?: {
    pickerLabel: string;
    getValue: () => string;
    subscribeToRangeChanges: (listener: () => void) => () => void;
  };
  strings: ConditionalFormatRuleDialogStrings;
}

export interface ConditionalFormatTextDialogOptions {
  title: string;
  label: string;
  initial?: string;
  initialStyle?: string;
  strings: ConditionalFormatRuleDialogStrings;
}

const appendFormatRow = (
  body: HTMLElement,
  strings: ConditionalFormatRuleDialogStrings,
  initialStyle: string | undefined,
): {
  select: HTMLSelectElement;
  preview: HTMLElement;
  getStyle: () => ConditionalFormatDialogStyle;
} => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block fc-tb__dlg__format-choice';
  const label = document.createElement('label');
  label.className = 'fc-tb__dlg__label';
  label.textContent = strings.formatWith;
  const select = createDialogSelect(
    conditionalStyleOptions(strings),
    initialStyle ?? DEFAULT_CF_STYLE_ID,
    { className: 'fc-tb__dlg__select' },
  );
  let customStyle: ConditionalFormatDialogStyle | null = null;
  const preview = document.createElement('span');
  preview.className = 'fc-tb__dlg__format-preview';
  preview.dataset.conditionalFormatPreview = 'true';
  preview.textContent = 'AaBbCcYyZz';
  preview.title = strings.formatPreview;
  const currentStyle = (): ConditionalFormatDialogStyle =>
    select.value === 'custom' && customStyle
      ? customStyle
      : conditionalStyleFromValue(select.value, strings);
  const updatePreview = (): void => {
    const selected = select.selectedOptions[0]?.textContent ?? strings.formatPreview;
    applyConditionalStylePreview(preview, currentStyle(), selected);
  };
  select.addEventListener('change', () => {
    updatePreview();
    if (select.value !== 'custom') return;
    void showConditionalFormatCustomStyleDialog(strings, currentStyle()).then((style) => {
      if (!style) return;
      customStyle = style;
      updatePreview();
    });
  });
  updatePreview();
  label.appendChild(select);
  label.appendChild(preview);
  row.appendChild(label);
  body.appendChild(row);
  return { select, preview, getStyle: currentStyle };
};

const readNumber = (
  input: HTMLInputElement,
  opts: { locale?: string; min?: number; max?: number },
): number | null => {
  const value = parseLocalizedNumericInput(input.value, opts.locale);
  if (value === null) return null;
  if (typeof opts.min === 'number' && value < opts.min) return null;
  if (typeof opts.max === 'number' && value > opts.max) return null;
  return value;
};

const formatLocalizedNumericInitial = (
  value: number | undefined,
  locale: string | undefined,
): string => {
  if (!Number.isFinite(value)) return '';
  try {
    return new Intl.NumberFormat(locale, {
      useGrouping: false,
      maximumSignificantDigits: 17,
    }).format(value as number);
  } catch {
    // A custom dictionary may use an invalid BCP 47 tag. Keep the dialog
    // usable while the input parser falls back to its default grammar.
    return String(value);
  }
};

export const showConditionalFormatNumberDialog = (
  opts: ConditionalFormatNumberDialogOptions,
): Promise<ConditionalFormatDialogResult | null> =>
  new Promise<ConditionalFormatDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    const first = appendInputRow(shell.body, opts.label, {
      type: 'text',
      initial: formatLocalizedNumericInitial(opts.initial, opts.locale),
      min: opts.min,
      max: opts.max,
      step: opts.step,
    });
    first.inputMode = 'decimal';
    const second =
      opts.secondLabel !== undefined
        ? appendInputRow(shell.body, opts.secondLabel, {
            type: 'text',
            initial: formatLocalizedNumericInitial(opts.secondInitial, opts.locale),
            min: opts.min,
            max: opts.max,
            step: opts.step,
          })
        : null;
    if (second) second.inputMode = 'decimal';
    const { reference } = opts;
    if (reference) {
      for (const input of second ? [first, second] : [first]) {
        attachRangePickerButton(input, {
          label: reference.pickerLabel,
          getValue: reference.getValue,
          subscribeToRangeChanges: reference.subscribeToRangeChanges,
          kind: 'conditional-format-quick-value',
        });
      }
    }
    const readField = (
      input: HTMLInputElement,
    ): { value: number; reference: string | null } | null => {
      const text = input.value.trim();
      if (reference && text.length > 1 && text.startsWith('='))
        return { value: 0, reference: text };
      const value = readNumber(input, opts);
      return value === null ? null : { value, reference: null };
    };
    const formatRow = appendFormatRow(shell.body, opts.strings, opts.initialStyle);
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.strings.cancel,
      okLabel: opts.strings.ok,
    });

    const lifecycle = installDialogLifecycle<ConditionalFormatDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      const a = readField(first);
      const b = second ? readField(second) : null;
      if (a === null || (second && b === null)) {
        showInputError(
          errorRow,
          a === null ? first : (second ?? first),
          opts.strings.invalidNumber,
        );
        return;
      }
      const fields = b ? [a, b] : [a];
      lifecycle.finish({
        values: fields.map((field) => field.value),
        ...(reference ? { references: fields.map((field) => field.reference) } : {}),
        style: formatRow.getStyle(),
      });
    };

    mountDialog(shell, () => focusAndSelectInput(first));
  });

export const showConditionalFormatTextDialog = (
  opts: ConditionalFormatTextDialogOptions,
): Promise<ConditionalFormatTextDialogResult | null> =>
  new Promise<ConditionalFormatTextDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    const input = appendInputRow(shell.body, opts.label, { initial: opts.initial ?? '' });
    const formatRow = appendFormatRow(shell.body, opts.strings, opts.initialStyle);
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.strings.cancel,
      okLabel: opts.strings.ok,
    });

    const lifecycle = installDialogLifecycle<ConditionalFormatTextDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      const text = input.value.trim();
      if (!text) {
        showInputError(errorRow, input, opts.strings.invalidText);
        return;
      }
      lifecycle.finish({ text, style: formatRow.getStyle() });
    };

    mountDialog(shell, () => focusAndSelectInput(input));
  });
