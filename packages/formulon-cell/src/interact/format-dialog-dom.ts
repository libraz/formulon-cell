import { type ColorPaletteHandle, createColorPalette } from '../components/color-palette.js';
import type { SideKey } from './format-dialog-model.js';

export function makeCheckbox(label: string): {
  wrap: HTMLLabelElement;
  input: HTMLInputElement;
} {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__check';
  const input = document.createElement('input');
  input.type = 'checkbox';
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return { wrap, input };
}

export function makeButton(label: string, primary = false): HTMLButtonElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = primary ? 'fc-fmtdlg__btn fc-fmtdlg__btn--primary' : 'fc-fmtdlg__btn';
  b.textContent = label;
  return b;
}

export function makeSwatches(
  kind: 'font' | 'border' | 'fill',
  themeLabel: string,
  standardLabel: string,
): ColorPaletteHandle {
  const palette = createColorPalette({
    themeLabel,
    standardLabel,
    ariaLabel: themeLabel,
    // Picks are committed by the dialog controller via click delegation on
    // the palette's `[data-color]` swatches; the palette only owns its own
    // highlight state here.
    onPick: () => {},
  });
  palette.el.dataset.swatches = kind;
  return palette;
}

export function makeVisualSideButton(
  visualSideButtons: Map<SideKey, HTMLButtonElement[]>,
  key: SideKey,
  label: string,
  extraClass = '',
): HTMLButtonElement {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = `fc-fmtdlg__border-hit fc-fmtdlg__border-hit--${key}${extraClass}`;
  btn.dataset.borderSide = key;
  btn.setAttribute('aria-label', label);
  btn.setAttribute('aria-pressed', 'false');
  const buttons = visualSideButtons.get(key) ?? [];
  buttons.push(btn);
  visualSideButtons.set(key, buttons);
  return btn;
}

export function makeSection(title: string): HTMLDivElement {
  const section = document.createElement('div');
  section.className = 'fc-fmtdlg__section';
  const heading = document.createElement('div');
  heading.className = 'fc-fmtdlg__section-title';
  heading.textContent = title;
  section.appendChild(heading);
  return section;
}

export function makeListSourceRadio(
  value: 'literal' | 'range',
  label: string,
): { wrap: HTMLLabelElement; input: HTMLInputElement } {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__check';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = 'fc-validation-list-source';
  input.value = value;
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return { wrap, input };
}
