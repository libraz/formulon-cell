// Color-scale subform of the Conditional Formatting dialog: two or three
// color stops, each with a scale-point threshold.

import type { Range } from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';
import type { ConditionalRule } from '../store/store.js';
import { appendScalePointRow, collectScalePoint } from './conditional-form-controls.js';

type ColorScaleRule = Extract<ConditionalRule, { kind: 'color-scale' }>;

export interface ColorScaleForm {
  readonly group: HTMLDivElement;
  readonly useThreeCk: HTMLInputElement;
  /** Show or hide the middle stop to match the three-stop checkbox. */
  syncThreeStops(): void;
  /** Re-run each threshold type's change handling (value visibility). */
  refreshScaleTypes(): void;
  reset(): void;
  populate(rule: ColorScaleRule): void;
  /** Null when a threshold value is not a number. */
  collect(range: Range): ColorScaleRule | null;
}

export function appendColorScaleForm(
  form: HTMLElement,
  t: Strings['conditionalDialog'],
): ColorScaleForm {
  const group = document.createElement('div');
  group.className = 'fc-conddlg__sub';
  form.appendChild(group);

  const useThreeRow = document.createElement('label');
  useThreeRow.className = 'fc-fmtdlg__check';
  const useThreeCk = document.createElement('input');
  useThreeCk.type = 'checkbox';
  const useThreeText = document.createElement('span');
  useThreeText.textContent = t.useThreeStops;
  useThreeRow.append(useThreeCk, useThreeText);
  group.appendChild(useThreeRow);

  const stopMinRow = document.createElement('label');
  stopMinRow.className = 'fc-fmtdlg__row';
  const stopMinLabel = document.createElement('span');
  stopMinLabel.textContent = t.stopMin;
  const stopMinInput = document.createElement('input');
  stopMinInput.type = 'color';
  stopMinInput.value = '#f8696b';
  stopMinInput.setAttribute('aria-label', t.stopMin);
  stopMinRow.append(stopMinLabel, stopMinInput);
  group.appendChild(stopMinRow);

  const stopMidRow = document.createElement('label');
  stopMidRow.className = 'fc-fmtdlg__row';
  const stopMidLabel = document.createElement('span');
  stopMidLabel.textContent = t.stopMid;
  const stopMidInput = document.createElement('input');
  stopMidInput.type = 'color';
  stopMidInput.value = '#ffeb84';
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
  stopMaxInput.value = '#63be7b';
  stopMaxInput.setAttribute('aria-label', t.stopMax);
  stopMaxRow.append(stopMaxLabel, stopMaxInput);
  group.appendChild(stopMaxRow);

  const scaleMin = appendScalePointRow(group, t, t.stopMin, 'min', '0');
  const scaleMid = appendScalePointRow(group, t, t.stopMid, 'percentile', '50');
  scaleMid.row.hidden = true;
  const scaleMax = appendScalePointRow(group, t, t.stopMax, 'max', '100');

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
    scaleMin.value.value = '0';
    scaleMax.type.value = 'max';
    scaleMax.value.value = '100';
    scaleMid.type.value = 'percentile';
    scaleMid.value.value = '50';
  };

  const populate = (rule: ColorScaleRule): void => {
    useThreeCk.checked = rule.stops.length === 3;
    stopMinInput.value = rule.stops[0] ?? '#f8696b';
    stopMidInput.value = rule.stops.length === 3 ? (rule.stops[1] ?? '#ffeb84') : '#ffeb84';
    stopMaxInput.value = rule.stops.at(-1) ?? '#63be7b';
    const thresholds = rule.thresholds ?? [];
    const min = thresholds[0];
    const mid = rule.stops.length === 3 ? thresholds[1] : undefined;
    const max = rule.stops.length === 3 ? thresholds[2] : thresholds[1];
    if (min) {
      scaleMin.type.value = min.kind;
      scaleMin.value.value = 'value' in min ? String(min.value) : '0';
    }
    if (mid) {
      scaleMid.type.value = mid.kind;
      scaleMid.value.value = 'value' in mid ? String(mid.value) : '50';
    }
    if (max) {
      scaleMax.type.value = max.kind;
      scaleMax.value.value = 'value' in max ? String(max.value) : '100';
    }
  };

  const collect = (range: Range): ColorScaleRule | null => {
    const stops: [string, string] | [string, string, string] = useThreeCk.checked
      ? [stopMinInput.value, stopMidInput.value, stopMaxInput.value]
      : [stopMinInput.value, stopMaxInput.value];
    const minPoint = collectScalePoint(scaleMin);
    const maxPoint = collectScalePoint(scaleMax);
    if (!minPoint || !maxPoint) return null;
    if (useThreeCk.checked) {
      const midPoint = collectScalePoint(scaleMid);
      if (!midPoint) return null;
      return { kind: 'color-scale', range, stops, thresholds: [minPoint, midPoint, maxPoint] };
    }
    return { kind: 'color-scale', range, stops, thresholds: [minPoint, maxPoint] };
  };

  return { group, useThreeCk, syncThreeStops, refreshScaleTypes, reset, populate, collect };
}
