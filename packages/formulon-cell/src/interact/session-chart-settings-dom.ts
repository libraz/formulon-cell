/** DOM helpers for the chart settings dialog: view-state capture and field factories. */

import { DEFAULT_THEME_ACCENTS } from '../charts/series-colors.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';

const SVG_SAFE_HEX = /^#[0-9a-f]{6}$/i;

const supportsTextSelection = (
  value: Element | null,
): value is HTMLInputElement | HTMLTextAreaElement =>
  value instanceof HTMLTextAreaElement ||
  (value instanceof HTMLInputElement &&
    ['password', 'search', 'tel', 'text', 'url'].includes(value.type));

interface SettingsViewState {
  focusedIndex: number;
  selectionStart: number | null;
  selectionEnd: number | null;
  selectionDirection: 'forward' | 'backward' | 'none' | null;
  surfaceScrollTop: number;
  bodyScrollTop: number;
}

export const captureSettingsView = (surface: HTMLElement): SettingsViewState => {
  const active = document.activeElement;
  const controls = Array.from(
    surface.querySelectorAll<HTMLElement>('button, input, select, textarea'),
  );
  const selectionTarget = supportsTextSelection(active) ? active : null;
  const body = surface.querySelector<HTMLElement>('.fc-chart-settings__body');
  return {
    focusedIndex: active instanceof HTMLElement ? controls.indexOf(active) : -1,
    selectionStart: selectionTarget?.selectionStart ?? null,
    selectionEnd: selectionTarget?.selectionEnd ?? null,
    selectionDirection: selectionTarget?.selectionDirection ?? null,
    surfaceScrollTop: surface.scrollTop,
    bodyScrollTop: body?.scrollTop ?? 0,
  };
};

export const restoreSettingsView = (surface: HTMLElement, state: SettingsViewState): void => {
  const controls = Array.from(
    surface.querySelectorAll<HTMLElement>('button, input, select, textarea'),
  );
  const focused = controls[state.focusedIndex];
  if (focused) {
    focused.focus({ preventScroll: true });
    if (
      supportsTextSelection(focused) &&
      state.selectionStart !== null &&
      state.selectionEnd !== null
    ) {
      focused.setSelectionRange(
        state.selectionStart,
        state.selectionEnd,
        state.selectionDirection ?? 'none',
      );
    }
  }
  surface.scrollTop = state.surfaceScrollTop;
  const body = surface.querySelector<HTMLElement>('.fc-chart-settings__body');
  if (body) body.scrollTop = state.bodyScrollTop;
};

export const element = <K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
): HTMLElementTagNameMap[K] => {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
};

export const textNode = (
  tag: 'span' | 'div' | 'h3' | 'legend',
  text: string,
  className?: string,
): HTMLElement => {
  const node = element(tag, className);
  node.textContent = text;
  return node;
};

const makeField = (parent: HTMLElement, label: string, control: HTMLElement): HTMLElement => {
  const row = element('label', 'fc-chart-settings__field');
  row.append(textNode('span', label, 'fc-chart-settings__label'), control);
  parent.appendChild(row);
  return row;
};

export const makeCheckbox = (
  parent: HTMLElement,
  label: string,
  checked: boolean,
  onChange: (checked: boolean) => void,
  disabled = false,
  className = '',
): HTMLInputElement => {
  const row = element('label', `fc-chart-settings__check ${className}`.trim());
  const input = element('input');
  input.type = 'checkbox';
  input.checked = checked;
  projectDisabledState(input, disabled, null);
  input.addEventListener('change', () => onChange(input.checked));
  row.append(input, textNode('span', label));
  parent.appendChild(row);
  return input;
};

interface TextInputOptions {
  readonly allowNegative?: boolean;
  readonly min?: number;
  readonly max?: number;
  readonly step?: string;
  readonly placeholder?: string;
  readonly disabled?: boolean;
}

export const makeTextInput = (
  parent: HTMLElement,
  label: string,
  value: string,
  onInput: (value: string) => void,
  type: 'text' | 'number' = 'text',
  options: TextInputOptions = {},
): HTMLInputElement => {
  const input = element('input');
  input.type = type;
  input.value = value;
  if (options.placeholder !== undefined) input.placeholder = options.placeholder;
  if (type === 'number') {
    input.inputMode = 'decimal';
    if (options.min !== undefined) input.min = String(options.min);
    else if (!options.allowNegative) input.min = '0';
    if (options.max !== undefined) input.max = String(options.max);
    if (options.step !== undefined) input.step = options.step;
  }
  projectDisabledState(input, options.disabled === true, null);
  input.addEventListener('input', () => onInput(input.value));
  makeField(parent, label, input);
  return input;
};

export const makeSelect = (
  parent: HTMLElement,
  label: string,
  value: string,
  options: readonly { value: string; label: string; disabled?: boolean }[],
  onChange: (value: string) => void,
  disabled = false,
): HTMLSelectElement => {
  const select = element('select');
  for (const option of options) {
    const item = element('option');
    item.value = option.value;
    item.textContent = option.label;
    projectDisabledState(item, option.disabled === true, null);
    select.appendChild(item);
  }
  projectDisabledState(select, disabled, null);
  select.value = value;
  select.addEventListener('change', () => onChange(select.value));
  makeField(parent, label, select);
  return select;
};

const validColor = (value: string): string | undefined =>
  SVG_SAFE_HEX.test(value.trim()) ? value.trim().toLowerCase() : undefined;

export const makeColorField = (
  parent: HTMLElement,
  label: string,
  value: string,
  onChange: (value: string) => void,
): void => {
  const row = element('div', 'fc-chart-settings__field fc-chart-settings__color-field');
  row.appendChild(textNode('span', label, 'fc-chart-settings__label'));
  const controls = element('span', 'fc-chart-settings__color-controls');
  const text = element('input');
  text.type = 'text';
  text.setAttribute('aria-label', label);
  text.inputMode = 'text';
  text.value = value;
  text.maxLength = 7;
  const native = element('input');
  native.type = 'color';
  native.setAttribute('aria-label', label);
  native.value = validColor(value) ?? DEFAULT_THEME_ACCENTS[0] ?? '#000000';
  const apply = (next: string): void => {
    const color = validColor(next);
    if (!color) return;
    text.value = color;
    native.value = color;
    onChange(color);
  };
  text.addEventListener('input', () => apply(text.value));
  native.addEventListener('input', () => apply(native.value));
  controls.append(text, native);
  row.appendChild(controls);
  parent.appendChild(row);
};
