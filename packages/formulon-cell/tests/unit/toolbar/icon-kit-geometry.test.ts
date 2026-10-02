import { describe, expect, it } from 'vitest';
import { sector } from '../../../src/toolbar/icon-kit/path.js';
import { badge } from '../../../src/toolbar/icon-kit/primitives.js';
import { PALETTE } from '../../../src/toolbar/icon-kit/tokens.js';

describe('toolbar/icon kit geometry', () => {
  it('builds circular sectors with explicit sweep and large-arc flags', () => {
    const clockwise = sector(12, 12, 8.8, -Math.PI / 2, 0);
    expect(clockwise).toContain('A8.8 8.8 0 0 1');
    expect(clockwise).toContain('L12 3.2');
    expect(clockwise).toContain('20.8 12z');

    const counterClockwise = sector(12, 12, 8.8, 0, -Math.PI / 2);
    expect(counterClockwise).toContain('A8.8 8.8 0 0 0');
    expect(sector(12, 12, 8.8, Math.PI / 4, (3 * Math.PI) / 2)).toContain('A8.8 8.8 0 1 1');
    expect(sector(12, 12, 8.8, 0, 2 * Math.PI)).toMatch(/^M3.2 12a8.8 8.8/);
    expect(sector(12, 12, 8.8, 0, 0)).toBe('');
    expect(() => sector(12, 12, 0, 0, 1)).toThrow(RangeError);
    expect(() => sector(12, 12, 8.8, 0, 2 * Math.PI + 0.01)).toThrow(RangeError);
    expect(() => sector(12, 12, Number.NaN, 0, 1)).toThrow(RangeError);
  });

  it('keeps badges inside an eight-unit plate and gives pencil badges a silhouette', () => {
    const segments = badge({ glyph: 'pencil', corner: 'tr', tone: PALETTE.alt });
    expect(segments[0]?.d).toMatch(/^M16.6 1/);

    const pencil = segments.slice(1);
    expect(pencil).toHaveLength(2);
    expect(pencil[0]?.fill).toBe(PALETTE.paper);
    const coordinates = [...(pencil[0]?.d.matchAll(/-?\d+(?:\.\d+)?/g) ?? [])].map((match) =>
      Number(match[0]),
    );
    const xs = coordinates.filter((_, index) => index % 2 === 0);
    const ys = coordinates.filter((_, index) => index % 2 === 1);
    expect(Math.min(...xs)).toBeGreaterThanOrEqual(15);
    expect(Math.max(...xs)).toBeLessThanOrEqual(23);
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(1);
    expect(Math.max(...ys)).toBeLessThanOrEqual(9);
    expect(pencil[1]?.stroke).toBe(PALETTE.alt);
    expect(pencil.every((segment) => !segment.strokeDasharray)).toBe(true);
  });
});
