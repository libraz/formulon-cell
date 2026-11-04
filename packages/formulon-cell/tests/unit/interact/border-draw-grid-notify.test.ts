import { describe, expect, it, vi } from 'vitest';
import { attachBorderDraw } from '../../../src/interact/border-draw.js';
import { cellRect, layoutForView } from '../../../src/render/geometry.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('border grid drag', () => {
  it('writes the dragged rectangle in one store update', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const pen = attachBorderDraw({ host, store });
    pen.activate('grid');
    const s = store.getState();
    const fire = (type: string, row: number, col: number): void => {
      const r = cellRect(layoutForView(s), s.viewport, row, col);
      host.dispatchEvent(
        new MouseEvent(type, {
          clientX: r.x + r.w / 2,
          clientY: r.y + r.h / 2,
          bubbles: true,
          cancelable: true,
        }),
      );
    };
    fire('pointerdown', 0, 0);
    const listener = vi.fn();
    store.subscribe(listener);
    fire('pointerup', 2, 2);
    expect(listener).toHaveBeenCalledTimes(1);
    expect(store.getState().format.formats.get('0:2:2')?.borders?.top).toBeTruthy();
    expect(store.getState().format.formats.size).toBe(9);
    pen.detach();
    document.body.innerHTML = '';
  });
});
