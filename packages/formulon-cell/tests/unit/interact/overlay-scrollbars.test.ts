import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { attachOverlayScrollbars, overlayThumb } from '../../../src/interact/overlay-scrollbars.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

describe('overlayThumb', () => {
  it('sizes the thumb by the visible share of the used extent or the view, whichever is longer', () => {
    // 199 used rows, 86 visible from row 7: thumb ≈ 86/199 of the track.
    const thumb = overlayThumb({ track: 1213, start: 7, visible: 86, used: 199 });
    expect(thumb.length).toBeCloseTo((1213 * 86) / 199, 6);
    expect(thumb.offset).toBeCloseTo((1213 * 7) / 199, 6);
    // A used range shorter than the view spans the whole track.
    expect(overlayThumb({ track: 800, start: 0, visible: 60, used: 12 })).toEqual({
      length: 800,
      offset: 0,
    });
  });
});

describe('attachOverlayScrollbars', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    vi.useFakeTimers();
    grid = document.createElement('div');
    Object.defineProperty(grid, 'clientHeight', { value: 600 });
    Object.defineProperty(grid, 'clientWidth', { value: 1000 });
    document.body.appendChild(grid);
    store = createSpreadsheetStore();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 198, col: 11 }), {
        value: { kind: 'number', value: 1 },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
  });

  afterEach(() => {
    vi.useRealTimers();
    grid.remove();
  });

  const thumbs = () => [...grid.querySelectorAll<HTMLElement>('.fc-scrollthumb')];

  it('shows the thumbs only while the view scrolls', () => {
    const handle = attachOverlayScrollbars({ grid, store });
    expect(thumbs()).toHaveLength(2);
    expect(thumbs().every((t) => t.dataset.visible !== 'true')).toBe(true);
    expect(thumbs().every((t) => t.getAttribute('aria-hidden') === 'true')).toBe(true);

    mutators.scrollBy(store, 5, 0);
    expect(thumbs().every((t) => t.dataset.visible === 'true')).toBe(true);
    const vertical = grid.querySelector<HTMLElement>('.fc-scrollthumb--v');
    expect(Number.parseFloat(vertical?.style.height ?? '0')).toBeGreaterThan(0);
    expect(Number.parseFloat(vertical?.style.top ?? '0')).toBeGreaterThan(0);

    vi.advanceTimersByTime(2000);
    expect(thumbs().every((t) => t.dataset.visible !== 'true')).toBe(true);
    handle.detach();
    expect(thumbs()).toHaveLength(0);
  });
});
