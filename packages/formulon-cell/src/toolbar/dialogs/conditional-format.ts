import {
  applyConditionalStylePreview,
  type ConditionalFormatDialogStyle,
  conditionalStyleFromValue,
  conditionalStyleOptions,
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

export type { ConditionalFormatDialogStyle } from './conditional-format-style.js';
export {
  applyConditionalStylePreview,
  conditionalStyleFromValue,
  conditionalStyleOptions,
} from './conditional-format-style.js';

export interface ConditionalFormatRuleDialogStrings {
  ok: string;
  cancel: string;
  formatWith: string;
  formatPreview: string;
  customFormat: string;
  customFormatTitle: string;
  customFillColor: string;
  customTextColor: string;
  customBold: string;
  customItalic: string;
  customUnderline: string;
  customStrike: string;
  formatLightRed: string;
  formatYellow: string;
  formatGreen: string;
  formatLightRedFill: string;
  formatRedText: string;
  formatRedBorder: string;
  formatRedFill: string;
  formatRedTextFill: string;
  invalidNumber: string;
  invalidText: string;
}

export interface ConditionalFormatDialogResult {
  readonly values: readonly number[];
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
    initialStyle ?? 'light-red-dark-red',
    { className: 'fc-tb__dlg__select' },
  );
  let customStyle: ConditionalFormatDialogStyle | null = null;
  const preview = document.createElement('span');
  preview.className = 'fc-tb__dlg__format-preview';
  preview.dataset.conditionalFormatPreview = 'true';
  preview.textContent = 'AaBbCcYyZz';
  preview.title = strings.formatPreview;
  const updatePreview = (): void => {
    const selected = select.selectedOptions[0]?.textContent ?? strings.formatPreview;
    applyConditionalStylePreview(
      preview,
      select.value === 'custom' && customStyle
        ? customStyle
        : conditionalStyleFromValue(select.value, strings),
      selected,
    );
  };
  select.addEventListener('change', () => {
    updatePreview();
    if (select.value !== 'custom') return;
    void showConditionalFormatCustomStyleDialog(
      strings,
      customStyle ?? conditionalStyleFromValue(select.value, strings),
    ).then((style) => {
      if (!style) return;
      customStyle = style;
      updatePreview();
    });
  });
  updatePreview();
  const getStyle = (): ConditionalFormatDialogStyle =>
    select.value === 'custom' && customStyle
      ? customStyle
      : conditionalStyleFromValue(select.value, strings);
  label.appendChild(select);
  label.appendChild(preview);
  row.appendChild(label);
  body.appendChild(row);
  return { select, preview, getStyle };
};

const readNumber = (
  input: HTMLInputElement,
  opts: { min?: number; max?: number },
): number | null => {
  const value = Number(input.value);
  if (!Number.isFinite(value)) return null;
  if (typeof opts.min === 'number' && value < opts.min) return null;
  if (typeof opts.max === 'number' && value > opts.max) return null;
  return value;
};

export const showConditionalFormatNumberDialog = (
  opts: ConditionalFormatNumberDialogOptions,
): Promise<ConditionalFormatDialogResult | null> =>
  new Promise<ConditionalFormatDialogResult | null>((resolve) => {
    const shell = createDialogShell({ title: opts.title });
    const first = appendInputRow(shell.body, opts.label, {
      type: 'number',
      initial: Number.isFinite(opts.initial) ? String(opts.initial) : '',
      min: opts.min,
      max: opts.max,
      step: opts.step,
    });
    const second =
      opts.secondLabel !== undefined
        ? appendInputRow(shell.body, opts.secondLabel, {
            type: 'number',
            initial: Number.isFinite(opts.secondInitial) ? String(opts.secondInitial) : '',
            min: opts.min,
            max: opts.max,
            step: opts.step,
          })
        : null;
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
    });
    const onOk = (): void => {
      const a = readNumber(first, opts);
      const b = second ? readNumber(second, opts) : null;
      if (a === null || (second && b === null)) {
        showInputError(
          errorRow,
          a === null ? first : (second ?? first),
          opts.strings.invalidNumber,
        );
        return;
      }
      lifecycle.finish({
        values: second ? [a, b as number] : [a],
        style: formatRow.getStyle(),
      });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

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
    });
    const onOk = (): void => {
      const text = input.value.trim();
      if (!text) {
        showInputError(errorRow, input, opts.strings.invalidText);
        return;
      }
      lifecycle.finish({ text, style: formatRow.getStyle() });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(input));
  });
