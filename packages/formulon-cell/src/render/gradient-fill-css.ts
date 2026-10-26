import type { CellGradientFill } from '../store/store.js';

/** CSS background for an OOXML gradient fill. OOXML measures a linear angle
 *  clockwise from left-to-right, CSS from bottom-to-top; a path gradient is
 *  approximated by a radial one around its inner rectangle. */
export const gradientFillCss = (gradient: CellGradientFill): string => {
  const stops = gradient.stops
    .map((stop) => `${stop.color} ${Math.round(Math.max(0, Math.min(1, stop.position)) * 100)}%`)
    .join(', ');
  if (gradient.type === 'path') {
    const x = Math.round(((gradient.left + gradient.right) / 2) * 100);
    const y = Math.round(((gradient.top + gradient.bottom) / 2) * 100);
    return `radial-gradient(farthest-corner at ${x}% ${y}%, ${stops})`;
  }
  return `linear-gradient(${gradient.degree + 90}deg, ${stops})`;
};
