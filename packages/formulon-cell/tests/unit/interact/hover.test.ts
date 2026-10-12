import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { attachHover } from '../../../src/interact/hover.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const tip = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-hover-tip');

const stubGridRect = (grid: HTMLElement): void => {
  grid.getBoundingClientRect = (): DOMRect =>
    ({ left: 0, top: 0, right: 800, bottom: 600, width: 800, height: 600, x: 0, y: 0 }) as DOMRect;
};

/** Uses the store's default layout so this helper stays aligned with the
 *  Excel-like sheet density. */
const cellPoint = (store: SpreadsheetStore, row: number, col: number): { x: number; y: number } => {
  const {
    headerColWidth: headerW,
    headerRowHeight: headerH,
    defaultColWidth: colW,
    defaultRowHeight: rowH,
  } = store.getState().layout;
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
    const { x, y } = cellPoint(store, 0, 0);
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

  it('pointermove over a hyperlink cell uses its authored tooltip when no comment exists', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { hyperlink: 'https://example.test/', hyperlinkTooltip: 'Open documentation' },
    );
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(store, 0, 0);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: x, clientY: y, pointerId: 1 }));

    expect(tip()?.hidden).toBe(false);
    expect(tip()?.textContent).toBe('Open documentation');
    handle.detach();
  });

  it('gives a cell comment priority over its hyperlink tooltip', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        comment: 'Review this cell',
        hyperlink: 'https://example.test/',
        hyperlinkTooltip: 'Open documentation',
      },
    );
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(store, 0, 0);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: x, clientY: y, pointerId: 1 }));

    expect(tip()?.textContent).toBe('Review this cell');
    handle.detach();
  });

  it('does not fall back to displaying the hyperlink URL as a tooltip', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { hyperlink: 'https://example.test/' },
    );
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(store, 0, 0);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: x, clientY: y, pointerId: 1 }));

    expect(tip()?.hidden).toBe(true);
    expect(tip()?.textContent).toBe('');
    handle.detach();
  });

  it('pointerleave hides the tooltip', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'note' });
    const handle = attachHover({ grid, store });
    const { x, y } = cellPoint(store, 0, 0);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: x, clientY: y, pointerId: 1 }));
    expect(tip()?.hidden).toBe(false);
    grid.dispatchEvent(new PointerEvent('pointerleave', { pointerId: 1 }));
    expect(tip()?.hidden).toBe(true);
    handle.detach();
  });

  it('pointermove over a cell without comment hides the tooltip', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'first' });
    const handle = attachHover({ grid, store });
    let p = cellPoint(store, 0, 0);
    grid.dispatchEvent(
      new PointerEvent('pointermove', { clientX: p.x, clientY: p.y, pointerId: 1 }),
    );
    expect(tip()?.hidden).toBe(false);
    // Hover a different cell with no format.
    p = cellPoint(store, 2, 2);
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
    const { x, y } = cellPoint(store, 1, 1);
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

  it('does not open unsafe hyperlink targets or advertise them as clickable', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      { hyperlink: 'javascript:alert(1)' },
    );
    const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
    const handle = attachHover({ grid, store });
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Control', ctrlKey: true }));
    const { x, y } = cellPoint(store, 1, 1);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: x, clientY: y }));
    expect(grid.style.cursor).toBe('');
    for (const target of ['javascript:alert(1)', 'data:text/html,<script>alert(1)</script>']) {
      mutators.setCellFormat(store, { sheet: 0, row: 1, col: 1 }, { hyperlink: target });
      const evt = new MouseEvent('click', {
        clientX: x,
        clientY: y,
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      grid.dispatchEvent(evt);
      expect(evt.defaultPrevented).toBe(false);
    }
    expect(openSpy).not.toHaveBeenCalled();
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
    const { x, y } = cellPoint(store, 1, 1);
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

  it('restores the cursor when leaving a link or releasing its modifier', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { hyperlink: 'https://example.test/' },
    );
    const handle = attachHover({ grid, store });
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Control', ctrlKey: true }));
    const link = cellPoint(store, 0, 0);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: link.x, clientY: link.y }));
    expect(grid.style.cursor).toBe('pointer');
    const plain = cellPoint(store, 0, 1);
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: plain.x, clientY: plain.y }));
    expect(grid.style.cursor).toBe('');
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: link.x, clientY: link.y }));
    expect(grid.style.cursor).toBe('pointer');
    window.dispatchEvent(new KeyboardEvent('keyup', { key: 'Control' }));
    expect(grid.style.cursor).toBe('');
    handle.detach();
  });

  it('keeps cell hover notes as compact rectangular Excel-style tips', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/panels/hover-tip.css'), 'utf8');

    expect(css).toMatch(
      /\.fc-hover-tip\s*\{[\s\S]*?box-shadow:[\s\S]*?0 4px 12px rgba\(0, 0, 0, 0\.14\),[\s\S]*?padding: 5px 8px;[\s\S]*?border-radius: 2px;/,
    );
    expect(css).not.toContain('box-shadow: var(--fc-shadow-4)');
  });
});
