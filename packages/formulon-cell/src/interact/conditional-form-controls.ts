// Form controls shared by the Conditional Formatting dialog's rule subforms:
// the plain select factory and the scale-point (type + value) rows used by
// color scales, data bars and icon-set thresholds.

import type { Strings } from '../i18n/strings.js';
import type { ConditionalScalePoint } from '../store/store.js';
import { createDialogSelect, type DialogSelectOption } from '../toolbar/dialogs/form-controls.js';

type ConditionalStrings = Strings['conditionalDialog'];

export interface ScalePointControl {
  type: HTMLSelectElement;
  value: HTMLInputElement;
  /** Present on icon-set thresholds: `>=` (default) or strict `>`. */
  operator?: HTMLSelectElement;
}

export const conditionalSelect = (
  options: readonly DialogSelectOption[],
  initial = options[0]?.value ?? '',
): HTMLSelectElement => createDialogSelect(options, initial, { className: '' });

export const scaleTypeOptions = (t: ConditionalStrings) =>
  [
    { id: 'min', label: t.scaleTypeMin },
    { id: 'max', label: t.scaleTypeMax },
    { id: 'number', label: t.scaleTypeNumber },
    { id: 'percent', label: t.scaleTypePercent },
    { id: 'percentile', label: t.scaleTypePercentile },
  ] as const;

/** Min/max points carry no value, so their value input hides itself. */
export const bindScalePointValueVisibility = (
  type: HTMLSelectElement,
  value: HTMLInputElement,
): void => {
  const syncValue = (): void => {
    value.hidden = type.value === 'min' || type.value === 'max';
  };
  type.addEventListener('change', syncValue);
  syncValue();
};

export function appendScalePointRow(
  parent: HTMLElement,
  t: ConditionalStrings,
  label: string,
  defaultType: ConditionalScalePoint['kind'],
  defaultValue: string,
  excludedKinds: readonly ConditionalScalePoint['kind'][] = [],
): { row: HTMLLabelElement; type: HTMLSelectElement; value: HTMLInputElement } {
  const row = document.createElement('label');
  row.className = 'fc-fmtdlg__row';
  const span = document.createElement('span');
  span.textContent = `${label} ${t.scaleType}`;
  const type = conditionalSelect(
    scaleTypeOptions(t)
      .filter((option) => !excludedKinds.includes(option.id))
      .map((option) => ({ value: option.id, label: option.label })),
    defaultType,
  );
  type.setAttribute('aria-label', `${label} ${t.scaleType}`);
  const value = document.createElement('input');
  value.type = 'number';
  value.value = defaultValue;
  value.setAttribute('aria-label', `${label} ${t.scaleValue}`);
  bindScalePointValueVisibility(type, value);
  row.append(span, type, value);
  parent.appendChild(row);
  return { row, type, value };
}

export const collectScalePoint = (input: ScalePointControl): ConditionalScalePoint | null => {
  const kind = input.type.value as ConditionalScalePoint['kind'];
  const comparison = input.operator?.value === '>' ? { gte: false } : {};
  if (kind === 'min' || kind === 'max') return { kind, ...comparison };
  const value = Number.parseFloat(input.value.value);
  if (!Number.isFinite(value)) return null;
  return { kind, value, ...comparison };
};

export const applyScalePoint = (
  control: ScalePointControl,
  point: ConditionalScalePoint | undefined,
  fallback: ConditionalScalePoint,
  fallbackValue = '0',
): void => {
  const next = point ?? fallback;
  control.type.value = next.kind;
  control.value.value = 'value' in next ? String(next.value) : fallbackValue;
  if (control.operator) control.operator.value = next.gte === false ? '>' : '>=';
  control.type.dispatchEvent(new Event('change'));
};
