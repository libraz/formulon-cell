import { describe, expect, it } from 'vitest';
import {
  clamp,
  clampPanelToViewport,
  panelSize,
  viewportSize,
} from '../../../src/interact/overlay-position.js';

describe('interact/overlay-position', () => {
  it('reports viewport size with document fallback', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 640 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 360 });

    expect(viewportSize()).toEqual({ width: 640, height: 360 });
  });

  it('clamps with a non-inverted range even when the panel is wider than the viewport', () => {
    expect(clamp(100, 8, -40)).toBe(8);
    expect(clamp(-100, 8, 200)).toBe(8);
    expect(clamp(120, 8, 200)).toBe(120);
    expect(clamp(999, 8, 200)).toBe(200);
  });

  it('measures panels from rect, offset, or fallback dimensions', () => {
    const panel = document.createElement('div');
    panel.getBoundingClientRect = () =>
      ({ width: 123.2, height: 45.1, left: 0, top: 0, right: 123.2, bottom: 45.1 }) as DOMRect;
    expect(panelSize(panel, 10, 10)).toEqual({ width: 124, height: 46 });

    panel.getBoundingClientRect = () =>
      ({ width: 0, height: 0, left: 0, top: 0, right: 0, bottom: 0 }) as DOMRect;
    Object.defineProperty(panel, 'offsetWidth', { configurable: true, value: 90 });
    Object.defineProperty(panel, 'offsetHeight', { configurable: true, value: 80 });
    expect(panelSize(panel, 10, 10)).toEqual({ width: 90, height: 80 });

    Object.defineProperty(panel, 'offsetWidth', { configurable: true, value: 0 });
    Object.defineProperty(panel, 'offsetHeight', { configurable: true, value: 0 });
    expect(panelSize(panel, 10, 12)).toEqual({ width: 10, height: 12 });
  });

  it('clamps panel positions using configurable padding and fallback size', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 180 });
    const panel = document.createElement('div');

    expect(
      clampPanelToViewport(panel, 310, 170, {
        pad: 4,
        fallbackWidth: 292,
        fallbackHeight: 160,
      }),
    ).toEqual({ x: 24, y: 16 });
  });
});
