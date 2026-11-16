import { afterEach, describe, expect, it, vi } from 'vitest';

const layoutCalls = vi.hoisted(() => ({ count: 0 }));

vi.mock('../../../src/render/geometry.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../../src/render/geometry.js')>();
  return {
    ...actual,
    layoutForView: (...args: Parameters<typeof actual.layoutForView>) => {
      layoutCalls.count += 1;
      return actual.layoutForView(...args);
    },
  };
});

import { attachHover } from '../../../src/interact/hover.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('attachHover measurement per move', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('reads the grid rect and builds the layout once for a pointer move', () => {
    const grid = document.createElement('div');
    grid.dataset.fcPlatform = 'mac';
    document.body.appendChild(grid);
    let rectReads = 0;
    grid.getBoundingClientRect = (): DOMRect => {
      rectReads += 1;
      return { left: 0, top: 0, right: 800, bottom: 600, width: 800, height: 600 } as DOMRect;
    };
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'note' });
    const handle = attachHover({
      grid,
      store,
      staleTip: {
        strings: () => {
          throw new Error('not a stale cell');
        },
        isStale: () => false,
      },
    });
    const { headerColWidth: hx, headerRowHeight: hy } = store.getState().layout;
    const move = (): void => {
      grid.dispatchEvent(
        new PointerEvent('pointermove', { clientX: hx + 5, clientY: hy + 5, bubbles: true }),
      );
    };
    move();
    expect(document.querySelector('.fc-hover-tip--note')).not.toBeNull();
    rectReads = 0;
    layoutCalls.count = 0;
    move();
    expect({ rectReads, layoutBuilds: layoutCalls.count }).toEqual({
      rectReads: 1,
      layoutBuilds: 1,
    });
    handle.detach();
  });
});
