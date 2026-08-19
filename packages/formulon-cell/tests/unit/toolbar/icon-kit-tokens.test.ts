import { describe, expect, it } from 'vitest';
import { EXCEL_RIBBON_ICON_PATHS } from '../../../src/toolbar/excel-ribbon-icons.js';
import { PALETTE, STROKE } from '../../../src/toolbar/icon-kit/tokens.js';

/**
 * The ribbon icons are composed from the icon kit rather than hand-authored, so
 * the properties the kit is supposed to hold constant are checkable here.
 * Geometry (optical size, centring, bleed) needs a real renderer and is covered
 * by `scripts/icons/render.mjs`.
 */

const entries = Object.entries(EXCEL_RIBBON_ICON_PATHS);
const allowedStrokeWidths = new Set<string>(Object.values(STROKE));
const allowedColors = new Set<string>([...Object.values(PALETTE), 'none']);

describe('toolbar/icon kit conformance', () => {
  it('defines every icon with at least one segment', () => {
    expect(entries.length).toBeGreaterThan(0);
    const empty = entries.filter(([, segments]) => segments.length === 0).map(([name]) => name);
    expect(empty).toEqual([]);
  });

  it('draws strokes only at the five widths on the scale', () => {
    const offScale = new Map<string, Set<string>>();
    for (const [name, segments] of entries) {
      for (const segment of segments) {
        if (segment.strokeWidth && !allowedStrokeWidths.has(segment.strokeWidth)) {
          const widths = offScale.get(name) ?? new Set<string>();
          widths.add(segment.strokeWidth);
          offScale.set(name, widths);
        }
      }
    }
    expect(Object.fromEntries([...offScale].map(([k, v]) => [k, [...v]]))).toEqual({});
  });

  it('paints only palette colours', () => {
    const offPalette = new Map<string, Set<string>>();
    for (const [name, segments] of entries) {
      for (const color of [segments].flat().flatMap((s) => [s.fill, s.stroke])) {
        if (color && !allowedColors.has(color)) {
          const colors = offPalette.get(name) ?? new Set<string>();
          colors.add(color);
          offPalette.set(name, colors);
        }
      }
    }
    expect(Object.fromEntries([...offPalette].map(([k, v]) => [k, [...v]]))).toEqual({});
  });

  it('emits well-formed path data', () => {
    const malformed = entries
      .flatMap(([name, segments]) => segments.map((segment) => [name, segment.d] as const))
      .filter(([, d]) => d.length === 0 || /NaN|Infinity|undefined/.test(d))
      .map(([name]) => name);
    expect([...new Set(malformed)]).toEqual([]);
  });

  it('keeps every stroked segment paired with a colour', () => {
    const unpainted = entries
      .flatMap(([name, segments]) => segments.map((segment) => [name, segment] as const))
      .filter(([, segment]) => Boolean(segment.strokeWidth) && !segment.stroke)
      .map(([name]) => name);
    expect([...new Set(unpainted)]).toEqual([]);
  });
});
