// Icon-set subform of the Conditional Formatting dialog: the icon family,
// order/visibility flags and up to four threshold rows (one fewer than the
// family's icon count is shown). An edited rule keeps its hydrated floor.

import type { Range } from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';
import type { ConditionalIconSet, ConditionalRule, ConditionalScalePoint } from '../store/store.js';
import type { DialogSelectOption } from '../toolbar/dialogs/form-controls.js';
import {
  bindScalePointValueVisibility,
  collectScalePoint,
  conditionalSelect,
  scaleTypeOptions,
} from './conditional-form-controls.js';

type IconSetRule = Extract<ConditionalRule, { kind: 'icon-set' }>;

export interface IconSetForm {
  readonly group: HTMLDivElement;
  readonly select: HTMLSelectElement;
  labelFor(id: ConditionalIconSet): string;
  /** Show the threshold rows the selected family uses; seed empty values. */
  syncThresholds(): void;
  /** Re-run each threshold type's change handling (value visibility). */
  refreshThresholdTypes(): void;
  reset(): void;
  populate(rule: IconSetRule): void;
  /** Null when a visible threshold value is not a number. */
  collect(range: Range): IconSetRule | null;
}

/** `isEditing` reports whether the dialog is editing an existing rule; only
 *  then is the hydrated floor written back. */
export function appendIconSetForm(
  form: HTMLElement,
  t: Strings['conditionalDialog'],
  isEditing: () => boolean,
): IconSetForm {
  const group = document.createElement('div');
  group.className = 'fc-conddlg__sub';
  form.appendChild(group);

  const iconSetRow = document.createElement('label');
  iconSetRow.className = 'fc-fmtdlg__row';
  const iconSetLabel = document.createElement('span');
  iconSetLabel.textContent = t.kindIconSet;
  const iconSetOptions: { id: ConditionalIconSet; label: string }[] = [
    { id: 'arrows3', label: t.iconSetArrows3 },
    { id: 'arrows5', label: t.iconSetArrows5 },
    { id: 'triangles3', label: t.iconSetTriangles3 },
    { id: 'traffic3', label: t.iconSetTraffic3 },
    { id: 'trafficRim3', label: t.iconSetTrafficRim3 },
    { id: 'symbols3', label: t.iconSetSymbols3 },
    { id: 'flags3', label: t.iconSetFlags3 },
    { id: 'stars3', label: t.iconSetStars3 },
    { id: 'quarters5', label: t.iconSetQuarters5 },
    { id: 'ratings5', label: t.iconSetRatings5 },
    { id: 'bars5', label: t.iconSetBars5 },
    { id: 'boxes5', label: t.iconSetBoxes5 },
  ];
  const iconSetSelect = conditionalSelect(
    iconSetOptions.map((o) => ({ value: o.id, label: o.label })),
  );
  const iconSetLabelFor = (id: ConditionalIconSet): string =>
    iconSetOptions.find((option) => option.id === id)?.label ?? id;
  iconSetRow.append(iconSetLabel, iconSetSelect);
  group.appendChild(iconSetRow);

  const iconReverseRow = document.createElement('label');
  iconReverseRow.className = 'fc-fmtdlg__check';
  const iconReverseCk = document.createElement('input');
  iconReverseCk.type = 'checkbox';
  const iconReverseText = document.createElement('span');
  iconReverseText.textContent = t.reverseOrder;
  iconReverseRow.append(iconReverseCk, iconReverseText);
  group.appendChild(iconReverseRow);

  const iconOnlyRow = document.createElement('label');
  iconOnlyRow.className = 'fc-fmtdlg__check';
  const iconOnlyCk = document.createElement('input');
  iconOnlyCk.type = 'checkbox';
  const iconOnlyText = document.createElement('span');
  iconOnlyText.textContent = t.showIconOnly;
  iconOnlyRow.append(iconOnlyCk, iconOnlyText);
  group.appendChild(iconOnlyRow);

  const iconOperatorOptions: DialogSelectOption[] = [
    { value: '>=', label: '>=' },
    { value: '>', label: '>' },
  ];

  const makeIconThresholdRow = (
    index: number,
  ): {
    row: HTMLLabelElement;
    operator: HTMLSelectElement;
    type: HTMLSelectElement;
    value: HTMLInputElement;
  } => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row fc-conddlg__icon-threshold-row';
    const span = document.createElement('span');
    span.textContent = `${t.iconThreshold} ${index + 1}`;
    const operator = conditionalSelect(iconOperatorOptions, '>=');
    operator.setAttribute('aria-label', `${t.iconThreshold} ${index + 1} ${t.iconOperator}`);
    operator.setAttribute('data-cf-icon-operator', String(index));
    const type = conditionalSelect(
      scaleTypeOptions(t).map((option) => ({ value: option.id, label: option.label })),
      'percent',
    );
    type.setAttribute('aria-label', `${t.iconThreshold} ${index + 1} ${t.scaleType}`);
    type.setAttribute('data-cf-icon-type', String(index));
    const value = document.createElement('input');
    value.type = 'number';
    value.setAttribute('aria-label', `${t.iconThreshold} ${index + 1} ${t.scaleValue}`);
    value.setAttribute('data-cf-icon-value', String(index));
    bindScalePointValueVisibility(type, value);
    row.append(span, operator, type, value);
    group.appendChild(row);
    return { row, operator, type, value };
  };
  const iconThresholdControls = [0, 1, 2, 3].map((index) => makeIconThresholdRow(index));

  let preservedIconFloor: ConditionalScalePoint | undefined;

  const syncIconThresholds = (): void => {
    const slots = iconSetSelect.value.endsWith('5') ? 5 : 3;
    for (let index = 0; index < iconThresholdControls.length; index += 1) {
      const control = iconThresholdControls[index];
      if (!control) continue;
      control.row.hidden = index >= slots - 1;
      if (control.value.value === '') {
        control.value.value = String(Math.round(((index + 1) * 100) / slots));
      }
    }
  };

  const refreshThresholdTypes = (): void => {
    for (const control of iconThresholdControls) control.type.dispatchEvent(new Event('change'));
  };

  const reset = (): void => {
    preservedIconFloor = undefined;
    iconSetSelect.value = 'arrows3';
    iconReverseCk.checked = false;
    iconOnlyCk.checked = false;
    for (const control of iconThresholdControls) {
      control.operator.value = '>=';
      control.type.value = 'percent';
      control.value.value = '';
      control.type.dispatchEvent(new Event('change'));
    }
  };

  const populate = (rule: IconSetRule): void => {
    preservedIconFloor = rule.floor;
    iconSetSelect.value = rule.icons;
    iconReverseCk.checked = rule.reverseOrder === true;
    iconOnlyCk.checked = rule.showValue === false;
    for (const [index, point] of (rule.thresholds ?? []).entries()) {
      const control = iconThresholdControls[index];
      if (!control) continue;
      control.operator.value = point.gte === false ? '>' : '>=';
      control.type.value = point.kind;
      control.value.value = 'value' in point ? String(point.value) : '';
    }
  };

  const collect = (range: Range): IconSetRule | null => {
    const iconThresholds = iconThresholdControls
      .filter((control) => !control.row.hidden)
      .map((control) => collectScalePoint(control));
    if (iconThresholds.some((point) => point === null)) return null;
    return {
      kind: 'icon-set',
      range,
      icons: iconSetSelect.value as ConditionalIconSet,
      showValue: !iconOnlyCk.checked,
      thresholds: iconThresholds as ConditionalScalePoint[],
      reverseOrder: iconReverseCk.checked,
      ...(isEditing() && preservedIconFloor ? { floor: preservedIconFloor } : {}),
    };
  };

  return {
    group,
    select: iconSetSelect,
    labelFor: iconSetLabelFor,
    syncThresholds: syncIconThresholds,
    refreshThresholdTypes,
    reset,
    populate,
    collect,
  };
}
