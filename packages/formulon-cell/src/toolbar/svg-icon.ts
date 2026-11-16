import { SVG_NS } from '../svg-namespace.js';

/** Build a decorative `currentColor` SVG whose body is one `<path>` per entry of `paths`. */
export const svgPathIcon = (
  paths: string | readonly string[],
  viewBox: string,
  className: string,
): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.classList.add(className);
  svg.setAttribute('viewBox', viewBox);
  svg.setAttribute('fill', 'currentColor');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  for (const d of typeof paths === 'string' ? [paths] : paths) {
    const path = document.createElementNS(SVG_NS, 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  }
  return svg;
};
