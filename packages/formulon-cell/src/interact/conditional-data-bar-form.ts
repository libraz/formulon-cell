// Data-bar subform of the Conditional Formatting dialog: bounds, direction,
// fill style and the bar/border/axis appearance. Colors and the axis/border
// choices remember their hydrated originals, so editing a rule writes back
// only what the user changed and keeps alpha the native color input drops.

import type { Range } from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';
import type { ConditionalRule } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendScalePointRow,
  applyScalePoint,
  collectScalePoint,
  conditionalSelect,
} from './conditional-form-controls.js';

type DataBarRule = Extract<ConditionalRule, { kind: 'data-bar' }>;
type DataBarAxisPosition = NonNullable<DataBarRule['axisPosition']>;

interface DataBarColorDraft {
  original: string | undefined;
  changed: boolean;
}

const DATA_BAR_COLOR_DEFAULTS = {
  positive: '#638ec6',
  negative: '#ff0000',
  border: '#638ec6',
  negativeBorder: '#ff0000',
  axis: '#000000',
} as const;

const cssColorComponentToByte = (raw: string): number | null => {
  const trimmed = raw.trim();
  const percent = trimmed.endsWith('%');
  const number = Number.parseFloat(percent ? trimmed.slice(0, -1) : trimmed);
  if (!Number.isFinite(number)) return null;
  const byte = percent ? (number * 255) / 100 : number;
  return Math.round(Math.max(0, Math.min(255, byte)));
};

/** Native color inputs accept opaque hex only. Keep alpha in the draft's
 * original CSS string and use this conversion only for the visible value. */
const cssColorToHex = (color: string | undefined, fallback: string): string => {
  const value = color?.trim() ?? '';
  const hex = value.match(/^#([0-9a-f]{3,8})$/i)?.[1];
  if (hex && (hex.length === 3 || hex.length === 4 || hex.length === 6 || hex.length === 8)) {
    const rgb = hex.length <= 4 ? hex.slice(0, 3) : hex.slice(0, 6);
    const expanded =
      rgb.length === 3
        ? rgb
            .split('')
            .map((part) => `${part}${part}`)
            .join('')
        : rgb;
    return `#${expanded.toLowerCase()}`;
  }
  const functionBody = value.match(/^rgba?\((.*)\)$/i)?.[1];
  if (functionBody) {
    const channels = functionBody
      .replaceAll('/', ' ')
      .split(/[\s,]+/)
      .filter(Boolean);
    const red = cssColorComponentToByte(channels[0] ?? '');
    const green = cssColorComponentToByte(channels[1] ?? '');
    const blue = cssColorComponentToByte(channels[2] ?? '');
    if (red !== null && green !== null && blue !== null) {
      return `#${[red, green, blue].map((part) => part.toString(16).padStart(2, '0')).join('')}`;
    }
  }
  return fallback;
};

export interface DataBarForm {
  readonly group: HTMLDivElement;
  /** Restore the defaults a new rule starts from. */
  reset(): void;
  populate(rule: DataBarRule): void;
  /** Null when a bound value is not a number. */
  collect(range: Range): DataBarRule | null;
}

/** `isEditing` reports whether the dialog is editing an existing rule, where
 *  unchanged optional appearance fields stay absent instead of defaulted. */
export function appendDataBarForm(
  form: HTMLElement,
  t: Strings['conditionalDialog'],
  isEditing: () => boolean,
): DataBarForm {
  const group = document.createElement('div');
  group.className = 'fc-conddlg__sub';
  form.appendChild(group);

  const dataBarMin = appendScalePointRow(group, t, t.barMin, 'min', '0', ['max']);
  dataBarMin.type.setAttribute('data-cf-bar-min-type', '');
  dataBarMin.value.setAttribute('data-cf-bar-min-value', '');
  const dataBarMax = appendScalePointRow(group, t, t.barMax, 'max', '100', ['min']);
  dataBarMax.type.setAttribute('data-cf-bar-max-type', '');
  dataBarMax.value.setAttribute('data-cf-bar-max-value', '');

  const barDirectionRow = document.createElement('label');
  barDirectionRow.className = 'fc-fmtdlg__row';
  const barDirectionLabel = document.createElement('span');
  barDirectionLabel.textContent = t.barDirection;
  const barDirectionSelect = conditionalSelect([
    { value: 'context', label: t.barDirectionContext },
    { value: 'left-to-right', label: t.barDirectionLeftToRight },
    { value: 'right-to-left', label: t.barDirectionRightToLeft },
  ]);
  barDirectionSelect.setAttribute('aria-label', t.barDirection);
  barDirectionSelect.setAttribute('data-cf-bar-direction', '');
  barDirectionRow.append(barDirectionLabel, barDirectionSelect);
  group.appendChild(barDirectionRow);

  const barFillStyleRow = document.createElement('label');
  barFillStyleRow.className = 'fc-fmtdlg__row';
  const barFillStyleLabel = document.createElement('span');
  barFillStyleLabel.textContent = t.barFillStyle;
  const barFillStyleSelect = conditionalSelect([
    { value: 'gradient', label: t.gradientFill },
    { value: 'solid', label: t.solidFill },
  ]);
  barFillStyleRow.append(barFillStyleLabel, barFillStyleSelect);
  group.appendChild(barFillStyleRow);

  const barColorRow = document.createElement('label');
  barColorRow.className = 'fc-fmtdlg__row';
  const barColorLabel = document.createElement('span');
  barColorLabel.textContent = t.barColor;
  const barColorInput = document.createElement('input');
  barColorInput.type = 'color';
  barColorInput.value = DATA_BAR_COLOR_DEFAULTS.positive;
  barColorInput.setAttribute('aria-label', t.barColor);
  barColorInput.setAttribute('data-cf-bar-positive-color', '');
  barColorRow.append(barColorLabel, barColorInput);
  group.appendChild(barColorRow);

  const barNegativeColorRow = document.createElement('label');
  barNegativeColorRow.className = 'fc-fmtdlg__row';
  const barNegativeColorLabel = document.createElement('span');
  barNegativeColorLabel.textContent = t.barNegativeColor;
  const barNegativeColorInput = document.createElement('input');
  barNegativeColorInput.type = 'color';
  barNegativeColorInput.value = DATA_BAR_COLOR_DEFAULTS.negative;
  barNegativeColorInput.setAttribute('aria-label', t.barNegativeColor);
  barNegativeColorInput.setAttribute('data-cf-bar-negative-color', '');
  barNegativeColorRow.append(barNegativeColorLabel, barNegativeColorInput);
  group.appendChild(barNegativeColorRow);

  const barBorderStyleRow = document.createElement('label');
  barBorderStyleRow.className = 'fc-fmtdlg__row';
  const barBorderStyleLabel = document.createElement('span');
  barBorderStyleLabel.textContent = t.barBorderStyle;
  const barBorderStyleSelect = conditionalSelect([
    { value: 'none', label: t.barBorderNone },
    { value: 'solid', label: t.barBorderSolid },
  ]);
  barBorderStyleSelect.setAttribute('aria-label', t.barBorderStyle);
  barBorderStyleSelect.setAttribute('data-cf-bar-border-style', '');
  barBorderStyleRow.append(barBorderStyleLabel, barBorderStyleSelect);
  group.appendChild(barBorderStyleRow);

  const barBorderColorRow = document.createElement('label');
  barBorderColorRow.className = 'fc-fmtdlg__row';
  const barBorderColorLabel = document.createElement('span');
  barBorderColorLabel.textContent = t.barBorderColor;
  const barBorderColorInput = document.createElement('input');
  barBorderColorInput.type = 'color';
  barBorderColorInput.value = DATA_BAR_COLOR_DEFAULTS.border;
  barBorderColorInput.setAttribute('aria-label', t.barBorderColor);
  barBorderColorInput.setAttribute('data-cf-bar-border-color', '');
  barBorderColorRow.append(barBorderColorLabel, barBorderColorInput);
  group.appendChild(barBorderColorRow);

  const barNegativeBorderColorRow = document.createElement('label');
  barNegativeBorderColorRow.className = 'fc-fmtdlg__row';
  const barNegativeBorderColorLabel = document.createElement('span');
  barNegativeBorderColorLabel.textContent = t.barNegativeBorderColor;
  const barNegativeBorderColorInput = document.createElement('input');
  barNegativeBorderColorInput.type = 'color';
  barNegativeBorderColorInput.value = DATA_BAR_COLOR_DEFAULTS.negativeBorder;
  barNegativeBorderColorInput.setAttribute('aria-label', t.barNegativeBorderColor);
  barNegativeBorderColorInput.setAttribute('data-cf-bar-negative-border-color', '');
  barNegativeBorderColorRow.append(barNegativeBorderColorLabel, barNegativeBorderColorInput);
  group.appendChild(barNegativeBorderColorRow);

  const barAxisPositionRow = document.createElement('label');
  barAxisPositionRow.className = 'fc-fmtdlg__row';
  const barAxisPositionLabel = document.createElement('span');
  barAxisPositionLabel.textContent = t.barAxisPosition;
  const barAxisPositionSelect = conditionalSelect([
    { value: 'automatic', label: t.barAxisAutomatic },
    { value: 'middle', label: t.barAxisMiddle },
    { value: 'none', label: t.barAxisNone },
  ]);
  barAxisPositionSelect.setAttribute('aria-label', t.barAxisPosition);
  barAxisPositionSelect.setAttribute('data-cf-bar-axis-position', '');
  barAxisPositionRow.append(barAxisPositionLabel, barAxisPositionSelect);
  group.appendChild(barAxisPositionRow);

  const barAxisColorRow = document.createElement('label');
  barAxisColorRow.className = 'fc-fmtdlg__row';
  const barAxisColorLabel = document.createElement('span');
  barAxisColorLabel.textContent = t.barAxisColor;
  const barAxisColorInput = document.createElement('input');
  barAxisColorInput.type = 'color';
  barAxisColorInput.value = DATA_BAR_COLOR_DEFAULTS.axis;
  barAxisColorInput.setAttribute('aria-label', t.barAxisColor);
  barAxisColorInput.setAttribute('data-cf-bar-axis-color', '');
  barAxisColorRow.append(barAxisColorLabel, barAxisColorInput);
  group.appendChild(barAxisColorRow);

  const showValueRow = document.createElement('label');
  showValueRow.className = 'fc-fmtdlg__check';
  const showValueCk = document.createElement('input');
  showValueCk.type = 'checkbox';
  showValueCk.checked = true;
  const showValueText = document.createElement('span');
  showValueText.textContent = t.showValue;
  showValueRow.append(showValueCk, showValueText);
  group.appendChild(showValueRow);

  const dataBarColorDrafts = {
    positive: { original: undefined, changed: false } as DataBarColorDraft,
    negative: { original: undefined, changed: false } as DataBarColorDraft,
    border: { original: undefined, changed: false } as DataBarColorDraft,
    negativeBorder: { original: undefined, changed: false } as DataBarColorDraft,
    axis: { original: undefined, changed: false } as DataBarColorDraft,
  };
  let dataBarAxisPositionOriginal: DataBarAxisPosition | undefined;
  let dataBarAxisPositionChanged = false;
  let dataBarBorderStyleChanged = false;
  const syncDataBarAppearance = (): void => {
    const borderVisible = barBorderStyleSelect.value === 'solid';
    projectDisabledState(barBorderColorInput, !borderVisible, null);
    projectDisabledState(barNegativeBorderColorInput, !borderVisible, null);
    projectDisabledState(barAxisColorInput, barAxisPositionSelect.value === 'none', null);
  };
  const markColorChanged = (draft: DataBarColorDraft): void => {
    draft.changed = true;
  };
  barColorInput.addEventListener('input', () => markColorChanged(dataBarColorDrafts.positive));
  barColorInput.addEventListener('change', () => markColorChanged(dataBarColorDrafts.positive));
  barNegativeColorInput.addEventListener('input', () =>
    markColorChanged(dataBarColorDrafts.negative),
  );
  barNegativeColorInput.addEventListener('change', () =>
    markColorChanged(dataBarColorDrafts.negative),
  );
  barBorderColorInput.addEventListener('input', () => markColorChanged(dataBarColorDrafts.border));
  barBorderColorInput.addEventListener('change', () => markColorChanged(dataBarColorDrafts.border));
  barNegativeBorderColorInput.addEventListener('input', () =>
    markColorChanged(dataBarColorDrafts.negativeBorder),
  );
  barNegativeBorderColorInput.addEventListener('change', () =>
    markColorChanged(dataBarColorDrafts.negativeBorder),
  );
  barAxisColorInput.addEventListener('input', () => markColorChanged(dataBarColorDrafts.axis));
  barAxisColorInput.addEventListener('change', () => markColorChanged(dataBarColorDrafts.axis));
  barBorderStyleSelect.addEventListener('change', () => {
    dataBarBorderStyleChanged = true;
    syncDataBarAppearance();
  });
  barAxisPositionSelect.addEventListener('change', () => {
    dataBarAxisPositionChanged = true;
    syncDataBarAppearance();
  });
  syncDataBarAppearance();

  const setDataBarColorInput = (
    input: HTMLInputElement,
    draft: DataBarColorDraft,
    value: string | undefined,
    fallback: string,
  ): void => {
    draft.original = value;
    draft.changed = false;
    input.value = cssColorToHex(value, fallback);
  };
  const resetDataBarAppearance = (): void => {
    setDataBarColorInput(
      barColorInput,
      dataBarColorDrafts.positive,
      undefined,
      DATA_BAR_COLOR_DEFAULTS.positive,
    );
    setDataBarColorInput(
      barNegativeColorInput,
      dataBarColorDrafts.negative,
      undefined,
      DATA_BAR_COLOR_DEFAULTS.negative,
    );
    setDataBarColorInput(
      barBorderColorInput,
      dataBarColorDrafts.border,
      undefined,
      DATA_BAR_COLOR_DEFAULTS.border,
    );
    setDataBarColorInput(
      barNegativeBorderColorInput,
      dataBarColorDrafts.negativeBorder,
      undefined,
      DATA_BAR_COLOR_DEFAULTS.negativeBorder,
    );
    setDataBarColorInput(
      barAxisColorInput,
      dataBarColorDrafts.axis,
      undefined,
      DATA_BAR_COLOR_DEFAULTS.axis,
    );
    dataBarAxisPositionOriginal = undefined;
    dataBarAxisPositionChanged = false;
    barAxisPositionSelect.value = 'automatic';
    dataBarBorderStyleChanged = false;
    barBorderStyleSelect.value = 'none';
    syncDataBarAppearance();
  };
  const collectDataBarColor = (
    input: HTMLInputElement,
    draft: DataBarColorDraft,
    fallback: string,
    required: boolean,
  ): string | undefined => {
    if (draft.changed) return input.value || fallback;
    if (draft.original !== undefined) return draft.original;
    if (!required && isEditing()) return undefined;
    return input.value || fallback;
  };

  const reset = (): void => {
    applyScalePoint(dataBarMin, undefined, { kind: 'min' }, '0');
    applyScalePoint(dataBarMax, undefined, { kind: 'max' }, '100');
    barFillStyleSelect.value = 'gradient';
    resetDataBarAppearance();
    showValueCk.checked = true;
    barDirectionSelect.value = 'context';
  };

  const populate = (rule: DataBarRule): void => {
    barFillStyleSelect.value = rule.gradient === false ? 'solid' : 'gradient';
    setDataBarColorInput(
      barColorInput,
      dataBarColorDrafts.positive,
      rule.color,
      DATA_BAR_COLOR_DEFAULTS.positive,
    );
    setDataBarColorInput(
      barNegativeColorInput,
      dataBarColorDrafts.negative,
      rule.negativeColor,
      DATA_BAR_COLOR_DEFAULTS.negative,
    );
    setDataBarColorInput(
      barBorderColorInput,
      dataBarColorDrafts.border,
      rule.borderColor,
      DATA_BAR_COLOR_DEFAULTS.border,
    );
    setDataBarColorInput(
      barNegativeBorderColorInput,
      dataBarColorDrafts.negativeBorder,
      rule.negativeBorderColor,
      DATA_BAR_COLOR_DEFAULTS.negativeBorder,
    );
    setDataBarColorInput(
      barAxisColorInput,
      dataBarColorDrafts.axis,
      rule.axisColor,
      DATA_BAR_COLOR_DEFAULTS.axis,
    );
    dataBarAxisPositionOriginal = rule.axisPosition;
    dataBarAxisPositionChanged = false;
    barAxisPositionSelect.value = rule.axisPosition ?? 'automatic';
    dataBarBorderStyleChanged = false;
    barBorderStyleSelect.value =
      rule.borderColor !== undefined || rule.negativeBorderColor !== undefined ? 'solid' : 'none';
    showValueCk.checked = rule.showValue !== false;
    applyScalePoint(dataBarMin, rule.min, { kind: 'min' }, '0');
    applyScalePoint(dataBarMax, rule.max, { kind: 'max' }, '100');
    barDirectionSelect.value = rule.direction ?? 'context';
    syncDataBarAppearance();
  };

  const collect = (range: Range): DataBarRule | null => {
    const min = collectScalePoint(dataBarMin);
    const max = collectScalePoint(dataBarMax);
    if (!min || !max) return null;
    const axisPosition = barAxisPositionSelect.value as DataBarAxisPosition;
    const dataBarRule: DataBarRule = {
      kind: 'data-bar',
      range,
      color:
        collectDataBarColor(
          barColorInput,
          dataBarColorDrafts.positive,
          DATA_BAR_COLOR_DEFAULTS.positive,
          true,
        ) ?? DATA_BAR_COLOR_DEFAULTS.positive,
      min,
      max,
      direction: barDirectionSelect.value as Extract<
        ConditionalRule,
        { kind: 'data-bar' }
      >['direction'],
      gradient: barFillStyleSelect.value === 'gradient',
      showValue: showValueCk.checked,
    };
    if (dataBarAxisPositionChanged || dataBarAxisPositionOriginal !== undefined || !isEditing()) {
      dataBarRule.axisPosition = axisPosition;
    }
    const negativeColor = collectDataBarColor(
      barNegativeColorInput,
      dataBarColorDrafts.negative,
      DATA_BAR_COLOR_DEFAULTS.negative,
      false,
    );
    if (negativeColor !== undefined) dataBarRule.negativeColor = negativeColor;
    if (barBorderStyleSelect.value === 'solid') {
      const writeBothBorderColors = dataBarBorderStyleChanged || !isEditing();
      const borderColor = collectDataBarColor(
        barBorderColorInput,
        dataBarColorDrafts.border,
        DATA_BAR_COLOR_DEFAULTS.border,
        writeBothBorderColors,
      );
      const negativeBorderColor = collectDataBarColor(
        barNegativeBorderColorInput,
        dataBarColorDrafts.negativeBorder,
        DATA_BAR_COLOR_DEFAULTS.negativeBorder,
        writeBothBorderColors,
      );
      if (borderColor !== undefined) dataBarRule.borderColor = borderColor;
      if (negativeBorderColor !== undefined) dataBarRule.negativeBorderColor = negativeBorderColor;
    }
    const axisColor = collectDataBarColor(
      barAxisColorInput,
      dataBarColorDrafts.axis,
      DATA_BAR_COLOR_DEFAULTS.axis,
      false,
    );
    if (axisColor !== undefined) dataBarRule.axisColor = axisColor;
    return dataBarRule;
  };

  return { group, reset, populate, collect };
}
