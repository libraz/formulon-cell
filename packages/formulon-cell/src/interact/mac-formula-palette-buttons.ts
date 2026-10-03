// Action buttons for the Mac formula palette: text buttons and the stroked
// 16px icon buttons (close, range picker) in the palette header and fields.

import { createInteractionButton } from './chip-button.js';

export const makeButton = (
  label: string,
  action: string,
  onClick: () => void,
): HTMLButtonElement => {
  const button = createInteractionButton({ className: '', text: label, dataset: { action } });
  button.addEventListener('click', onClick);
  return button;
};

type PaletteIcon = 'close' | 'range';

const makeIcon = (kind: PaletteIcon): SVGSVGElement => {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '14');
  svg.setAttribute('height', '14');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.4');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  if (kind === 'close') {
    const first = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    first.setAttribute('d', 'M4 4l8 8');
    svg.appendChild(first);
    const second = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    second.setAttribute('d', 'M12 4l-8 8');
    svg.appendChild(second);
  } else {
    const grid = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    grid.setAttribute('x', '2');
    grid.setAttribute('y', '2');
    grid.setAttribute('width', '8');
    grid.setAttribute('height', '8');
    svg.appendChild(grid);
    const arrow = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    arrow.setAttribute('d', 'M8 13h5V8');
    svg.appendChild(arrow);
    const arrowHead = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    arrowHead.setAttribute('d', 'M13 8l-3 3');
    svg.appendChild(arrowHead);
  }
  return svg;
};

export const makeIconButton = (
  label: string,
  action: string,
  kind: PaletteIcon,
  onClick: () => void,
): HTMLButtonElement => {
  const button = createInteractionButton({ className: '', ariaLabel: label, dataset: { action } });
  button.title = label;
  button.appendChild(makeIcon(kind));
  button.addEventListener('click', onClick);
  return button;
};
