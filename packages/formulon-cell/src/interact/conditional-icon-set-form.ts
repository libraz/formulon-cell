// Icon-set subform of the Conditional Formatting dialog: the icon family,
// order/visibility flags, one icon picker per level and up to four threshold
// rows (one fewer than the family's icon count is shown). An edited rule
// keeps its hydrated floor.

import type { Range } from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';
import { iconSetSlotCount } from '../render/conditional.js';
import { paintConditionalIcon } from '../render/painters/cell.js';
import type {
  ConditionalIconRef,
  ConditionalIconSet,
  ConditionalRule,
  ConditionalScalePoint,
} from '../store/store.js';
import type { DialogSelectOption } from '../toolbar/dialogs/form-controls.js';
import {
  bindScalePointValueVisibility,
  collectScalePoint,
  conditionalSelect,
  localizedScalePointValue,
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
  getLocale: () => string = () => 'en-US',
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
    { id: 'arrowsGray3', label: t.iconSetArrowsGray3 },
    { id: 'triangles3', label: t.iconSetTriangles3 },
    { id: 'arrows4', label: t.iconSetArrows4 },
    { id: 'arrowsGray4', label: t.iconSetArrowsGray4 },
    { id: 'arrows5', label: t.iconSetArrows5 },
    { id: 'arrowsGray5', label: t.iconSetArrowsGray5 },
    { id: 'traffic3', label: t.iconSetTraffic3 },
    { id: 'trafficRim3', label: t.iconSetTrafficRim3 },
    { id: 'signs3', label: t.iconSetSigns3 },
    { id: 'traffic4', label: t.iconSetTraffic4 },
    { id: 'redToBlack4', label: t.iconSetRedToBlack4 },
    { id: 'symbolsCircled3', label: t.iconSetSymbolsCircled3 },
    { id: 'symbols3', label: t.iconSetSymbols3 },
    { id: 'flags3', label: t.iconSetFlags3 },
    { id: 'stars3', label: t.iconSetStars3 },
    { id: 'ratings4', label: t.iconSetRatings4 },
    { id: 'quarters5', label: t.iconSetQuarters5 },
    { id: 'ratings5', label: t.iconSetRatings5 },
    { id: 'boxes5', label: t.iconSetBoxes5 },
  ];
  const iconSetSelect = conditionalSelect(
    iconSetOptions.map((o) => ({ value: o.id, label: o.label })),
  );
  const iconSetLabelFor = (id: ConditionalIconSet): string =>
    iconSetOptions.find((option) => option.id === id)?.label ?? id;
  iconSetSelect.setAttribute('data-cf-icon-family', '');
  iconSetRow.append(iconSetLabel, iconSetSelect);
  group.appendChild(iconSetRow);

  // One picker per level, highest first like the desktop dialog; the value is
  // `set:slot`, or `none` for a level that draws no icon.
  const NO_ICON = 'none';
  const iconValue = (icon: ConditionalIconRef | null): string =>
    icon ? `${icon.set}:${icon.slot}` : NO_ICON;
  const iconFromValue = (value: string): ConditionalIconRef | null => {
    const [set, slot] = value.split(':');
    return set && slot !== undefined
      ? { set: set as ConditionalIconSet, slot: Number(slot) }
      : null;
  };
  const iconChoices = [
    { value: NO_ICON, label: t.noCellIcon },
    ...iconSetOptions.flatMap((option) =>
      Array.from({ length: iconSetSlotCount(option.id) }, (_, slot) => ({
        value: iconValue({ set: option.id, slot }),
        label: `${option.label} ${slot + 1}`,
      })),
    ),
  ];
  const drawPreview = (canvas: HTMLCanvasElement, value: string): void => {
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const icon = iconFromValue(value);
    if (icon) paintConditionalIcon(ctx, { x: 0, y: 0, w: 16, h: 16 }, icon.set, icon.slot);
  };
  const iconLevelControls = [4, 3, 2, 1, 0].map((level) => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row fc-conddlg__icon-level-row';
    const span = document.createElement('span');
    span.textContent = `${t.iconLevel} ${level + 1}`;
    const preview = document.createElement('canvas');
    preview.width = 16;
    preview.height = 16;
    preview.className = 'fc-conddlg__icon-preview';
    preview.setAttribute('aria-hidden', 'true');
    const select = conditionalSelect(iconChoices);
    select.setAttribute('aria-label', `${t.iconLevel} ${level + 1}`);
    select.setAttribute('data-cf-icon-level', String(level));
    select.addEventListener('change', () => drawPreview(preview, select.value));
    row.append(span, preview, select);
    group.appendChild(row);
    return { level, row, select, preview };
  });
  const setLevelIcon = (level: number, icon: ConditionalIconRef | null): void => {
    const control = iconLevelControls.find((entry) => entry.level === level);
    if (!control) return;
    control.select.value = iconValue(icon);
    drawPreview(control.preview, control.select.value);
  };
  // Family the pickers were last seeded from; a family change reseeds them.
  let pickerFamily: ConditionalIconSet | undefined;

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
    const family = iconSetSelect.value as ConditionalIconSet;
    const slots = iconSetSlotCount(family);
    if (pickerFamily !== family) {
      pickerFamily = family;
      for (let level = 0; level < 5; level += 1) {
        setLevelIcon(level, level < slots ? { set: family, slot: level } : null);
      }
    }
    for (const control of iconLevelControls) control.row.hidden = control.level >= slots;
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
    pickerFamily = undefined;
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
    pickerFamily = rule.icons;
    for (let level = 0; level < iconSetSlotCount(rule.icons); level += 1) {
      setLevelIcon(
        level,
        rule.customIcons ? (rule.customIcons[level] ?? null) : { set: rule.icons, slot: level },
      );
    }
    iconReverseCk.checked = rule.reverseOrder === true;
    iconOnlyCk.checked = rule.showValue === false;
    for (const [index, point] of (rule.thresholds ?? []).entries()) {
      const control = iconThresholdControls[index];
      if (!control) continue;
      control.operator.value = point.gte === false ? '>' : '>=';
      control.type.value = point.kind;
      control.value.value =
        'value' in point ? localizedScalePointValue(point.value, getLocale()) : '';
    }
  };

  const collect = (range: Range): IconSetRule | null => {
    const locale = getLocale();
    const iconThresholds = iconThresholdControls
      .filter((control) => !control.row.hidden)
      .map((control) => collectScalePoint(control, locale));
    if (iconThresholds.some((point) => point === null)) return null;
    const icons = iconSetSelect.value as ConditionalIconSet;
    const levels = Array.from({ length: iconSetSlotCount(icons) }, (_, level) =>
      iconFromValue(
        iconLevelControls.find((control) => control.level === level)?.select.value ?? NO_ICON,
      ),
    );
    const custom = levels.some((icon, level) => icon?.set !== icons || icon.slot !== level);
    return {
      kind: 'icon-set',
      range,
      icons,
      ...(custom ? { customIcons: levels } : {}),
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
