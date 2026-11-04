import { afterEach, describe, expect, it, vi } from 'vitest';
import { trackFrameDrag } from '../../../src/interact/drawing-frame.js';

const ptr = (type: string, x: number, y: number): PointerEvent => {
  const e = new Event(type, { cancelable: true }) as unknown as PointerEvent;
  Object.assign(e, { clientX: x, clientY: y, altKey: false });
  return e;
};

describe('trackFrameDrag pointercancel', () => {
  afterEach(() => vi.restoreAllMocks());

  const start = () => {
    const preview = vi.fn();
    const commit = vi.fn();
    trackFrameDrag(ptr('pointerdown', 0, 0), {
      start: { x: 0, y: 0, w: 100, h: 100 },
      min: { w: 10, h: 10 },
      keepRatio: () => false,
      preview,
      commit,
    });
    return { preview, commit };
  };

  it('detaches listeners without committing', () => {
    const { preview, commit } = start();
    window.dispatchEvent(ptr('pointermove', 50, 50));
    expect(preview).toHaveBeenCalledTimes(1);
    window.dispatchEvent(ptr('pointercancel', 50, 50));
    window.dispatchEvent(ptr('pointermove', 80, 80));
    window.dispatchEvent(ptr('pointerup', 80, 80));
    expect(preview).toHaveBeenCalledTimes(1);
    expect(commit).not.toHaveBeenCalled();
  });

  it('still commits on pointerup', () => {
    const { commit } = start();
    window.dispatchEvent(ptr('pointermove', 50, 50));
    window.dispatchEvent(ptr('pointerup', 50, 50));
    expect(commit).toHaveBeenCalledTimes(1);
  });
});
