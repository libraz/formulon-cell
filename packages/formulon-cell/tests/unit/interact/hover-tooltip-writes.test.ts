import { afterEach, describe, expect, it } from 'vitest';
import { attachHover } from '../../../src/interact/hover.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('attachHover tooltip writes', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('rewrites the tooltip text only when it changes', () => {
    const grid = document.createElement('div');
    document.body.appendChild(grid);
    grid.getBoundingClientRect = (): DOMRect =>
      ({ left: 0, top: 0, right: 800, bottom: 600, width: 800, height: 600 }) as DOMRect;
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'note' });
    const handle = attachHover({ grid, store });
    const el = document.querySelector<HTMLElement>('.fc-hover-tip') as HTMLElement;
    const { headerColWidth: hx, headerRowHeight: hy } = store.getState().layout;
    const move = (dx: number): void => {
      grid.dispatchEvent(
        new PointerEvent('pointermove', { clientX: hx + 5 + dx, clientY: hy + 5, bubbles: true }),
      );
    };
    move(0);
    expect(el.textContent).toBe('note');
    const observer = new MutationObserver(() => {});
    observer.observe(el, { childList: true, characterData: true, subtree: true, attributes: true });
    move(1);
    move(2);
    const records = observer.takeRecords();
    expect(records.filter((r) => r.attributeName !== 'style')).toEqual([]);
    expect(el.hidden).toBe(false);
    observer.disconnect();
    handle.detach();
  });
});
