import { describe, expect, it } from 'vitest';
import { flippedRotation, sameAnchor } from '../../../src/engine/drawing-geometry.js';
import type { EngineDrawingAnchor } from '../../../src/engine/types.js';

const anchor = { kind: 'twoCell', fromRow: 1, fromCol: 2, toRow: 3, toCol: 4 };

describe('sameAnchor', () => {
  it('ignores key order and treats two absent anchors as equal', () => {
    const reordered = { toCol: 4, toRow: 3, fromCol: 2, fromRow: 1, kind: 'twoCell' };
    expect(
      sameAnchor(
        anchor as unknown as EngineDrawingAnchor,
        reordered as unknown as EngineDrawingAnchor,
      ),
    ).toBe(true);
    expect(sameAnchor(undefined, undefined)).toBe(true);
  });

  it('tells differing anchors and an absent anchor apart', () => {
    const moved = { ...anchor, toRow: 9 };
    expect(
      sameAnchor(anchor as unknown as EngineDrawingAnchor, moved as unknown as EngineDrawingAnchor),
    ).toBe(false);
    expect(sameAnchor(anchor as unknown as EngineDrawingAnchor, undefined)).toBe(false);
  });
});

describe('flippedRotation', () => {
  it('mirrors the rotation about the vertical and keeps zero at zero', () => {
    expect(flippedRotation(30)).toBe(330);
    expect(flippedRotation(330)).toBe(30);
    expect(flippedRotation(0)).toBe(0);
    expect(flippedRotation()).toBe(0);
    expect(flippedRotation(180)).toBe(180);
    expect(flippedRotation(-90)).toBe(90);
    expect(flippedRotation(400)).toBe(320);
  });
});
