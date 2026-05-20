// Text-orientation dropdown for the Home tab. The menu items render a small
// inline SVG glyph showing the orientation effect (counter-clockwise, vertical,
// rotate up/down, etc.) plus the localized label.

import type { ToolbarMenuText } from '@libraz/formulon-cell';

import { SVG_NS } from '../border-icons.js';
import { createMenu, menuPresetButton, menuSeparator } from './general.js';

export type TextOrientationGlyph = 'ccw' | 'cw' | 'vertical' | 'up' | 'down' | 'format';

const createTextOrientationIcon = (glyph: TextOrientationGlyph): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.classList.add('app__text-orientation-preview');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.2');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  const baseline = document.createElementNS(SVG_NS, 'line');
  baseline.setAttribute('x1', '2');
  baseline.setAttribute('y1', '13');
  baseline.setAttribute('x2', '14');
  baseline.setAttribute('y2', '13');
  svg.appendChild(baseline);
  if (glyph === 'ccw' || glyph === 'cw') {
    const angle = glyph === 'ccw' ? -35 : 35;
    const text = document.createElementNS(SVG_NS, 'text');
    text.setAttribute('x', '4');
    text.setAttribute('y', '11');
    text.setAttribute('transform', `rotate(${angle} 8 11)`);
    text.setAttribute('font-family', 'system-ui, sans-serif');
    text.setAttribute('font-size', '7');
    text.setAttribute('font-weight', '700');
    text.setAttribute('fill', 'currentColor');
    text.setAttribute('stroke', 'none');
    text.textContent = 'ab';
    svg.appendChild(text);
  } else if (glyph === 'vertical') {
    for (let i = 0; i < 3; i += 1) {
      const ch = document.createElementNS(SVG_NS, 'text');
      ch.setAttribute('x', '8');
      ch.setAttribute('y', String(4 + i * 3));
      ch.setAttribute('text-anchor', 'middle');
      ch.setAttribute('font-family', 'system-ui, sans-serif');
      ch.setAttribute('font-size', '3');
      ch.setAttribute('font-weight', '700');
      ch.setAttribute('fill', 'currentColor');
      ch.setAttribute('stroke', 'none');
      ch.textContent = 'a';
      svg.appendChild(ch);
    }
  } else if (glyph === 'up' || glyph === 'down') {
    const text = document.createElementNS(SVG_NS, 'text');
    text.setAttribute('x', '0');
    text.setAttribute('y', '0');
    const rotate = glyph === 'up' ? -90 : 90;
    text.setAttribute('transform', `translate(8 11) rotate(${rotate})`);
    text.setAttribute('text-anchor', 'middle');
    text.setAttribute('font-family', 'system-ui, sans-serif');
    text.setAttribute('font-size', '7');
    text.setAttribute('font-weight', '700');
    text.setAttribute('fill', 'currentColor');
    text.setAttribute('stroke', 'none');
    text.textContent = 'ab';
    svg.appendChild(text);
  } else if (glyph === 'format') {
    const grid = document.createElementNS(SVG_NS, 'rect');
    grid.setAttribute('x', '2.5');
    grid.setAttribute('y', '3.5');
    grid.setAttribute('width', '11');
    grid.setAttribute('height', '7');
    svg.appendChild(grid);
    const hLine = document.createElementNS(SVG_NS, 'line');
    hLine.setAttribute('x1', '2.5');
    hLine.setAttribute('y1', '7');
    hLine.setAttribute('x2', '13.5');
    hLine.setAttribute('y2', '7');
    svg.appendChild(hLine);
    const vLine = document.createElementNS(SVG_NS, 'line');
    vLine.setAttribute('x1', '8');
    vLine.setAttribute('y1', '3.5');
    vLine.setAttribute('x2', '8');
    vLine.setAttribute('y2', '10.5');
    svg.appendChild(vLine);
  }
  return svg;
};

const textOrientationMenuItem = (
  glyph: TextOrientationGlyph,
  label: string,
  value: string,
): HTMLButtonElement => {
  return menuPresetButton(label, 'textOrientation', value, createTextOrientationIcon(glyph));
};

export const createTextOrientationMenu = (t: ToolbarMenuText): HTMLDivElement => {
  const menu = createMenu('menu-text-orientation');
  menu.append(
    textOrientationMenuItem('ccw', t.orientationAngleCounterclockwise, 'ccw'),
    textOrientationMenuItem('cw', t.orientationAngleClockwise, 'cw'),
    textOrientationMenuItem('vertical', t.orientationVerticalText, 'vertical'),
    textOrientationMenuItem('up', t.orientationRotateTextUp, 'up'),
    textOrientationMenuItem('down', t.orientationRotateTextDown, 'down'),
    menuSeparator(),
    textOrientationMenuItem('format', t.orientationFormatAlignment, 'format'),
  );
  return menu;
};
