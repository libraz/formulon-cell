import { describe, expect, it } from 'vitest';
import { drawingClipPath } from '../../../src/engine/drawing-geometry.js';
import {
  type DrawingFrame,
  type DrawingHandle,
  resizedFrame,
  rotatedResizedFrame,
} from '../../../src/interact/drawing-frame.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const MIN = { w: 10, h: 10 };
const START: DrawingFrame = { x: 100, y: 100, w: 80, h: 40 };

/** Screen point of the local point (lx, ly) of `frame` turned by `deg` about its centre. */
const screen = (frame: DrawingFrame, lx: number, ly: number, deg: number): [number, number] => {
  const t = (deg * Math.PI) / 180;
  const [cx, cy] = [frame.x + frame.w / 2, frame.y + frame.h / 2];
  const [ox, oy] = [frame.x + lx - cx, frame.y + ly - cy];
  return [cx + ox * Math.cos(t) - oy * Math.sin(t), cy + ox * Math.sin(t) + oy * Math.cos(t)];
};

const near = (actual: [number, number], expected: [number, number]) => {
  expect(actual[0]).toBeCloseTo(expected[0], 6);
  expect(actual[1]).toBeCloseTo(expected[1], 6);
};

describe('resizing a rotated frame', () => {
  it('at 0 degrees is the unrotated resize', () => {
    for (const handle of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as DrawingHandle[]) {
      expect(rotatedResizedFrame(START, handle, 13, -7, false, MIN, 0)).toEqual(
        resizedFrame(START, handle, 13, -7, false, MIN),
      );
    }
  });

  it('at 90 degrees, the right handle dragged down on screen widens the frame and holds the left edge', () => {
    const frame = rotatedResizedFrame(START, 'e', 0, 20, false, MIN, 90);

    expect(frame.w).toBeCloseTo(100);
    expect(frame.h).toBeCloseTo(40);
    near(screen(frame, 0, 20, 90), screen(START, 0, 20, 90));
  });

  it('at 30 degrees, a corner follows the pointer and the opposite corner stays put', () => {
    const frame = rotatedResizedFrame(START, 'se', 20, 10, false, MIN, 30);
    const t = (30 * Math.PI) / 180;
    // The pointer delta in the frame's own axes.
    const local = [20 * Math.cos(t) + 10 * Math.sin(t), -20 * Math.sin(t) + 10 * Math.cos(t)];

    expect(frame.w).toBeCloseTo(80 + (local[0] ?? 0));
    expect(frame.h).toBeCloseTo(40 + (local[1] ?? 0));
    near(screen(frame, 0, 0, 30), screen(START, 0, 0, 30));
  });

  it('at 30 degrees, a west edge keeps the east edge middle fixed', () => {
    const frame = rotatedResizedFrame(START, 'w', -15, 5, false, MIN, 30);

    near(screen(frame, frame.w, frame.h / 2, 30), screen(START, START.w, START.h / 2, 30));
  });

  it('keeps the minimum size and the ratio on a rotated corner', () => {
    const small = rotatedResizedFrame(START, 'se', -500, -500, false, MIN, 30);
    expect(small.w).toBe(MIN.w);
    expect(small.h).toBe(MIN.h);
    near(screen(small, 0, 0, 30), screen(START, 0, 0, 30));

    const kept = rotatedResizedFrame(START, 'se', 40, 0, true, MIN, 60);
    expect(kept.w / kept.h).toBeCloseTo(START.w / START.h);
  });
});

describe('clip path of a rotated drawing', () => {
  const store = createSpreadsheetStore();
  const anchor = { fromRow: 0, fromCol: 0 } as Parameters<typeof drawingClipPath>[1];
  const position = { x: 0, y: 0, w: 200, h: 100 };

  it('is the screen-axis inset unrotated and a polygon in the frame when rotated', () => {
    const flat = drawingClipPath(store, anchor, position);
    expect(flat).toMatch(/^inset\(/);
    const [top, , , lead] = flat.slice(6, -1).split(' ').map(parseFloat) as number[];

    const turned = drawingClipPath(store, anchor, position, 90);
    expect(turned).toMatch(/^polygon\(/);
    // Turned back onto the screen, the polygon is the same visible rectangle.
    const points = turned
      .slice(8, -1)
      .split(',')
      .map((p) => p.trim().split(' ').map(parseFloat) as [number, number]);
    const back = points.map(([x, y]) => {
      const [ox, oy] = [x - position.w / 2, y - position.h / 2];
      return [position.w / 2 - oy, position.h / 2 + ox];
    });
    expect(Math.max(...back.map((p) => p[0] ?? 0))).toBeGreaterThan(position.w);
    expect(Math.min(...back.map((p) => p[0] ?? 0))).toBeCloseTo(lead ?? 0);
    expect(Math.min(...back.map((p) => p[1] ?? 0))).toBeCloseTo(top ?? 0);
  });
});
