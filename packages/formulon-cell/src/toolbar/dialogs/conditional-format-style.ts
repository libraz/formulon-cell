import { appendCheckboxRow } from './form-controls.js';
import {
  appendDialogActions,
  appendInputRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface ConditionalFormatStyleStrings {
  ok: string;
  cancel: string;
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
}

export type ConditionalFormatDialogStyle = {
  readonly fill?: string;
  readonly color?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
};

type MutableConditionalFormatDialogStyle = {
  -readonly [K in keyof ConditionalFormatDialogStyle]: ConditionalFormatDialogStyle[K];
};

export interface ConditionalFormatStyleOption {
  value: string;
  label: string;
  style?: ConditionalFormatDialogStyle;
  custom?: boolean;
}

/** Style id of the first option, which is also the fallback style. */
export const DEFAULT_CF_STYLE_ID = 'light-red-dark-red';

const STYLE_PRESETS: readonly {
  value: string;
  label: keyof ConditionalFormatStyleStrings;
  style?: ConditionalFormatDialogStyle;
  custom?: boolean;
}[] = [
  {
    value: DEFAULT_CF_STYLE_ID,
    label: 'formatLightRed',
    style: { fill: '#ffc7ce', color: '#9c0006' },
  },
  {
    value: 'yellow-dark-yellow',
    label: 'formatYellow',
    style: { fill: '#ffeb9c', color: '#9c6500' },
  },
  {
    value: 'green-dark-green',
    label: 'formatGreen',
    style: { fill: '#c6efce', color: '#006100' },
  },
  { value: 'light-red-fill', label: 'formatLightRedFill', style: { fill: '#ffc7ce' } },
  { value: 'red-text', label: 'formatRedText', style: { color: '#9c0006' } },
  {
    value: 'red-border',
    label: 'formatRedBorder',
    style: { color: '#9c0006', underline: true },
  },
  { value: 'red-fill', label: 'formatRedFill', style: { fill: '#ff0000' } },
  {
    value: 'red-text-fill',
    label: 'formatRedTextFill',
    style: { fill: '#ff0000', color: '#ffffff' },
  },
  { value: 'custom', label: 'customFormat', custom: true },
];

const PRESET_STYLE_BY_VALUE = new Map(STYLE_PRESETS.map((preset) => [preset.value, preset.style]));
const FALLBACK_STYLE = PRESET_STYLE_BY_VALUE.get(
  DEFAULT_CF_STYLE_ID,
) as ConditionalFormatDialogStyle;

export const conditionalStyleOptions = (
  strings: ConditionalFormatStyleStrings,
): ConditionalFormatStyleOption[] =>
  STYLE_PRESETS.map((preset) => ({
    value: preset.value,
    label: strings[preset.label],
    ...(preset.style ? { style: { ...preset.style } } : {}),
    ...(preset.custom ? { custom: true } : {}),
  }));

export const conditionalStyleFromValue = (
  value: string,
  _strings: ConditionalFormatStyleStrings,
): ConditionalFormatDialogStyle => ({ ...(PRESET_STYLE_BY_VALUE.get(value) ?? FALLBACK_STYLE) });

export const applyConditionalStylePreview = (
  preview: HTMLElement,
  style: ConditionalFormatDialogStyle,
  label: string,
): void => {
  preview.style.background = style.fill ?? 'transparent';
  preview.style.color = style.color ?? 'inherit';
  preview.style.fontWeight = style.bold ? '700' : '400';
  preview.style.fontStyle = style.italic ? 'italic' : 'normal';
  preview.style.textDecoration = [
    style.underline ? 'underline' : '',
    style.strike ? 'line-through' : '',
  ]
    .filter(Boolean)
    .join(' ');
  preview.setAttribute('aria-label', label);
};

export const showConditionalFormatCustomStyleDialog = (
  strings: ConditionalFormatStyleStrings,
  initial: ConditionalFormatDialogStyle,
): Promise<ConditionalFormatDialogStyle | null> =>
  new Promise<ConditionalFormatDialogStyle | null>((resolve) => {
    const shell = createDialogShell({ title: strings.customFormatTitle });
    const fill = appendInputRow(shell.body, strings.customFillColor, {
      initial: initial.fill ?? '',
      placeholder: '#ffc7ce',
    });
    const color = appendInputRow(shell.body, strings.customTextColor, {
      initial: initial.color ?? '',
      placeholder: '#9c0006',
    });
    const bold = appendCheckboxRow(shell.body, strings.customBold, initial.bold === true, 'bold');
    const italic = appendCheckboxRow(
      shell.body,
      strings.customItalic,
      initial.italic === true,
      'italic',
    );
    const underline = appendCheckboxRow(
      shell.body,
      strings.customUnderline,
      initial.underline === true,
      'underline',
    );
    const strike = appendCheckboxRow(
      shell.body,
      strings.customStrike,
      initial.strike === true,
      'strike',
    );
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });
    const lifecycle = installDialogLifecycle<ConditionalFormatDialogStyle | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      const style: MutableConditionalFormatDialogStyle = {};
      const fillValue = fill.value.trim();
      const colorValue = color.value.trim();
      if (fillValue) style.fill = fillValue;
      if (colorValue) style.color = colorValue;
      if (bold.checked) style.bold = true;
      if (italic.checked) style.italic = true;
      if (underline.checked) style.underline = true;
      if (strike.checked) style.strike = true;
      lifecycle.finish(style);
    };
    mountDialog(shell, fill);
  });
