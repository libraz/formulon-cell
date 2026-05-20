import { type ColorPaletteHandle, createColorPalette } from '../components/color-palette.js';
import { createDialogButton, createDialogToggleButton } from './dialog-shell.js';
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
  return createDialogButton({ label, variant: primary ? 'primary' : 'secondary' });
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
  const btn = createDialogToggleButton({
    label,
    baseClass: `fc-fmtdlg__border-hit fc-fmtdlg__border-hit--${key}`,
    extraClass: extraClass.trim() || undefined,
    datasetKey: 'borderSide',
    value: key,
  });
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
