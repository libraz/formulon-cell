// Color-scale subform of the Conditional Formatting dialog: two or three
// color stops, each with a scale-point threshold.

import type { Range } from '../engine/types.js';
import { formatEditableNumber } from '../engine/value.js';
import type { Strings } from '../i18n/strings.js';
import type { ConditionalRule } from '../store/store.js';
import {
  appendScalePointRow,
  collectScalePoint,
  setMacScaleFields,
} from './conditional-form-controls.js';
import { makeCheckbox } from './format-dialog-dom.js';

type ColorScaleRule = Extract<ConditionalRule, { kind: 'color-scale' }>;

export interface ColorScaleForm {
  readonly group: HTMLDivElement;
  readonly useThreeCk: HTMLInputElement;
  /** Show or hide the middle stop to match the three-stop checkbox. */
  syncThreeStops(): void;
  /** Re-run each threshold type's change handling (value visibility). */
  refreshScaleTypes(): void;
  /** Show the automatic bounds as disabled fields named in parentheses, as Excel for Mac does. */
  setMacLayout(on: boolean): void;
  reset(): void;
  populate(rule: ColorScaleRule): void;
  /** Null when a threshold value is not a number. */
  collect(range: Range): ColorScaleRule | null;
}

const DEFAULT_STOP = { min: '#f8696b', mid: '#ffeb84', max: '#63be7b' } as const;
const DEFAULT_THRESHOLD = { min: '0', mid: '50', max: '100' } as const;

export function appendColorScaleForm(
  form: HTMLElement,
  t: Strings['conditionalDialog'],
  getLocale: () => string = () => 'en-US',
): ColorScaleForm {
  const group = document.createElement('div');
  group.className = 'fc-conddlg__sub';
  form.appendChild(group);

  const { wrap: useThreeRow, input: useThreeCk } = makeCheckbox(t.useThreeStops);
  group.appendChild(useThreeRow);

  const stopMinRow = document.createElement('label');
  stopMinRow.className = 'fc-fmtdlg__row';
  const stopMinLabel = document.createElement('span');
  stopMinLabel.textContent = t.stopMin;
  const stopMinInput = document.createElement('input');
  stopMinInput.type = 'color';
  stopMinInput.value = DEFAULT_STOP.min;
  stopMinInput.setAttribute('aria-label', t.stopMin);
  stopMinRow.append(stopMinLabel, stopMinInput);
  group.appendChild(stopMinRow);

  const stopMidRow = document.createElement('label');
  stopMidRow.className = 'fc-fmtdlg__row';
  const stopMidLabel = document.createElement('span');
  stopMidLabel.textContent = t.stopMid;
  const stopMidInput = document.createElement('input');
  stopMidInput.type = 'color';
  stopMidInput.value = DEFAULT_STOP.mid;
  stopMidInput.setAttribute('aria-label', t.stopMid);
  stopMidRow.append(stopMidLabel, stopMidInput);
  stopMidRow.hidden = true;
  group.appendChild(stopMidRow);

  const stopMaxRow = document.createElement('label');
  stopMaxRow.className = 'fc-fmtdlg__row';
  const stopMaxLabel = document.createElement('span');
  stopMaxLabel.textContent = t.stopMax;
  const stopMaxInput = document.createElement('input');
  stopMaxInput.type = 'color';
  stopMaxInput.value = DEFAULT_STOP.max;
  stopMaxInput.setAttribute('aria-label', t.stopMax);
  stopMaxRow.append(stopMaxLabel, stopMaxInput);
  group.appendChild(stopMaxRow);

  const scaleMin = appendScalePointRow(group, t, t.stopMin, 'min', DEFAULT_THRESHOLD.min);
  const scaleMid = appendScalePointRow(group, t, t.stopMid, 'percentile', DEFAULT_THRESHOLD.mid);
  scaleMid.row.hidden = true;
  const scaleMax = appendScalePointRow(group, t, t.stopMax, 'max', DEFAULT_THRESHOLD.max);

  const syncThreeStops = (): void => {
    stopMidRow.hidden = !useThreeCk.checked;
    scaleMid.row.hidden = !useThreeCk.checked;
  };

  const refreshScaleTypes = (): void => {
    scaleMin.type.dispatchEvent(new Event('change'));
    scaleMid.type.dispatchEvent(new Event('change'));
    scaleMax.type.dispatchEvent(new Event('change'));
  };

  const reset = (): void => {
    useThreeCk.checked = false;
    scaleMin.type.value = 'min';
    scaleMin.value.value = DEFAULT_THRESHOLD.min;
    scaleMax.type.value = 'max';
    scaleMax.value.value = DEFAULT_THRESHOLD.max;
    scaleMid.type.value = 'percentile';
    scaleMid.value.value = DEFAULT_THRESHOLD.mid;
  };

  const populate = (rule: ColorScaleRule): void => {
    useThreeCk.checked = rule.stops.length === 3;
    stopMinInput.value = rule.stops[0] ?? DEFAULT_STOP.min;
    stopMidInput.value =
      rule.stops.length === 3 ? (rule.stops[1] ?? DEFAULT_STOP.mid) : DEFAULT_STOP.mid;
    stopMaxInput.value = rule.stops.at(-1) ?? DEFAULT_STOP.max;
    const thresholds = rule.thresholds ?? [];
    const min = thresholds[0];
    const mid = rule.stops.length === 3 ? thresholds[1] : undefined;
    const max = rule.stops.length === 3 ? thresholds[2] : thresholds[1];
    if (min) {
      scaleMin.type.value = min.kind;
      scaleMin.value.value =
        'value' in min ? formatEditableNumber(min.value, getLocale()) : DEFAULT_THRESHOLD.min;
    }
    if (mid) {
      scaleMid.type.value = mid.kind;
      scaleMid.value.value =
        'value' in mid ? formatEditableNumber(mid.value, getLocale()) : DEFAULT_THRESHOLD.mid;
    }
    if (max) {
      scaleMax.type.value = max.kind;
      scaleMax.value.value =
        'value' in max ? formatEditableNumber(max.value, getLocale()) : DEFAULT_THRESHOLD.max;
    }
  };

  const collect = (range: Range): ColorScaleRule | null => {
    const stops: [string, string] | [string, string, string] = useThreeCk.checked
      ? [stopMinInput.value, stopMidInput.value, stopMaxInput.value]
      : [stopMinInput.value, stopMaxInput.value];
    const locale = getLocale();
    const minPoint = collectScalePoint(scaleMin, locale);
    const maxPoint = collectScalePoint(scaleMax, locale);
    if (!minPoint || !maxPoint) return null;
    if (useThreeCk.checked) {
      const midPoint = collectScalePoint(scaleMid, locale);
      if (!midPoint) return null;
      return { kind: 'color-scale', range, stops, thresholds: [minPoint, midPoint, maxPoint] };
    }
    return { kind: 'color-scale', range, stops, thresholds: [minPoint, maxPoint] };
  };

  const setMacLayout = (on: boolean): void => setMacScaleFields(group, on);
  return {
    group,
    useThreeCk,
    syncThreeStops,
    refreshScaleTypes,
    setMacLayout,
    reset,
    populate,
    collect,
  };
}
