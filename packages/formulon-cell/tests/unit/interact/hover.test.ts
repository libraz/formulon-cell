import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { attachHover } from '../../../src/interact/hover.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const tip = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-hover-tip');

const stubGridRect = (grid: HTMLElement): void => {
  grid.getBoundingClientRect = (): DOMRect =>
    ({ left: 0, top: 0, right: 800, bottom: 600, width: 800, height: 600, x: 0, y: 0 }) as DOMRect;
};

/** Default layout puts cell (0,0) starting at x=46, y=22 (header sizes). Row
 *  height defaults to 20, col width 64 — so (60,30) lands on cell (0,0). */
const cellPoint = (row: number, col: number): { x: number; y: number } => {
  const headerW = 46;
  const headerH = 22;
  const colW = 64;
  const rowH = 20;
  return {
    x: headerW + col * colW + Math.floor(colW / 2),
    y: headerH + row * rowH + Math.floor(rowH / 2),
  };
};

describe('attachHover', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    grid = document.createElement('div');
    document.body.appendChild(grid);
    stubGridRect(grid);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('attach mounts a hidden tooltip element under document.body', () => {
    const handle = attachHover({ grid, store });
    expect(tip()).not.toBeNull();
    expect(tip()?.hidden).toBe(true);
    handle.detach();
  });

  it('pointermove over a comment cell reveals the tooltip with the comment text', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'a sticky note' });
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(0, 0);
    grid.dispatchEvent(
      new PointerEvent('pointermove', {
        clientX: x,
        clientY: y,
        bubbles: true,
        pointerId: 1,
      }),
    );
    expect(tip()?.hidden).toBe(false);
    expect(tip()?.textContent).toBe('a sticky note');
    handle.detach();
  });

  it('pointerleave hides the tooltip', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'note' });
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(0, 0);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: x, clientY: y, pointerId: 1 }));
    expect(tip()?.hidden).toBe(false);
    grid.dispatchEvent(new PointerEvent('pointerleave', { pointerId: 1 }));
    expect(tip()?.hidden).toBe(true);
    handle.detach();
  });

  it('pointermove over a cell without comment hides the tooltip', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'first' });
    const handle = attachHover({ grid, store });
    let p = cellPoint(0, 0);
    grid.dispatchEvent(
      new PointerEvent('pointermove', { clientX: p.x, clientY: p.y, pointerId: 1 }),
    );
    expect(tip()?.hidden).toBe(false);
    // Hover a different cell with no format.
    p = cellPoint(2, 2);
    grid.dispatchEvent(
      new PointerEvent('pointermove', { clientX: p.x, clientY: p.y, pointerId: 1 }),
    );
    expect(tip()?.hidden).toBe(true);
    handle.detach();
  });

  it('Ctrl+click on a hyperlink cell calls window.open and is preventDefault-ed', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      { hyperlink: 'https://example.test/' },
    );
    const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(1, 1);
    const evt = new MouseEvent('click', {
      clientX: x,
      clientY: y,
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    });
    grid.dispatchEvent(evt);
    expect(openSpy).toHaveBeenCalledTimes(1);
    expect(openSpy).toHaveBeenCalledWith('https://example.test/', '_blank', 'noopener,noreferrer');
    expect(evt.defaultPrevented).toBe(true);
    openSpy.mockRestore();
    handle.detach();
  });

  it('plain click without modifier on a hyperlink cell does NOT open the link', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      { hyperlink: 'https://example.test/' },
    );
    const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(1, 1);
    grid.dispatchEvent(
      new MouseEvent('click', { clientX: x, clientY: y, bubbles: true, cancelable: true }),
    );
    expect(openSpy).not.toHaveBeenCalled();
    openSpy.mockRestore();
    handle.detach();
  });

  it('detach removes the tooltip element from the DOM', () => {
    const handle = attachHover({ grid, store });
    expect(tip()).not.toBeNull();
    handle.detach();
    expect(tip()).toBeNull();
  });
});
