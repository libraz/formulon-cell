import { afterEach, describe, expect, it, vi } from 'vitest';
import { attachHover } from '../../../src/interact/hover.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('attachHover format lookup', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  const setup = () => {
    const grid = document.createElement('div');
    document.body.appendChild(grid);
    grid.getBoundingClientRect = (): DOMRect =>
      ({ left: 0, top: 0, right: 800, bottom: 600, width: 800, height: 600 }) as DOMRect;
    const store = createSpreadsheetStore();
    const { headerColWidth: hx, headerRowHeight: hy } = store.getState().layout;
    const move = (dx: number): void => {
      grid.dispatchEvent(
        new PointerEvent('pointermove', { clientX: hx + 5 + dx, clientY: hy + 5, bubbles: true }),
      );
    };
    return { grid, store, move };
  };

  it('looks the cell format up once for repeated moves inside one cell', () => {
    const { grid, store, move } = setup();
    const handle = attachHover({ grid, store });
    const get = vi.spyOn(store.getState().format.formats, 'get');
    move(0);
    const afterFirst = get.mock.calls.length;
    move(1);
    move(2);
    expect(get.mock.calls.length).toBe(afterFirst);
    handle.detach();
  });

  it('refreshes the lookup after the format state changes', () => {
    const { grid, store, move } = setup();
    const handle = attachHover({ grid, store });
    const tip = document.querySelector<HTMLElement>('.fc-hover-tip') as HTMLElement;
    move(0);
    expect(tip.hidden).toBe(true);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'note' });
    move(1);
    expect(tip.textContent).toBe('note');
    handle.detach();
  });
});
