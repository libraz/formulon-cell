import { describe, expect, it } from 'vitest';
import { MAC_GRID_CURSORS } from '../../../src/interact/grid-cursors.js';

const svgOf = (cursor: string): string => {
  const m = cursor.match(/^url\("data:image\/svg\+xml,([^"]+)"\)/);
  if (!m?.[1]) throw new Error(`not an SVG cursor: ${cursor}`);
  return decodeURIComponent(m[1]);
};

describe('Mac grid cursors', () => {
  it.each(['colHeader', 'rowHeader', 'colResize', 'rowResize'] as const)(
    '%s draws its shape as a path, so the pointer stays visible',
    (key) => {
      const svg = svgOf(MAC_GRID_CURSORS[key]);
      expect(svg).toMatch(/<path d="M[^"]+"[^>]*\/>/);
      // Path data left as bare text between tags renders nothing.
      expect(svg.replace(/<[^>]*>/g, '').trim()).toBe('');
    },
  );
});
