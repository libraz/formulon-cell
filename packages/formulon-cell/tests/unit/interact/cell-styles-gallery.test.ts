import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { CELL_STYLES } from '../../../src/commands/cell-styles.js';
import { addrKey } from '../../../src/engine/address.js';
import { attachCellStylesGallery } from '../../../src/interact/cell-styles-gallery.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const overlay = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-stylegallery');
const chips = (): HTMLButtonElement[] =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-stylegallery__chip'));

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('attachCellStylesGallery', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setRange(store, 0, 0, 0, 0);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('mounts a hidden overlay and reveals it on open', () => {
    const handle = attachCellStylesGallery({ host, store });
    expect(overlay()?.hidden).toBe(true);
    handle.open();
    expect(overlay()?.hidden).toBe(false);
    handle.detach();
  });

  it('renders one chip per CELL_STYLES entry', () => {
    const handle = attachCellStylesGallery({ host, store });
    handle.open();
    const list = chips();
    expect(list.length).toBe(CELL_STYLES.length);
    expect(list.map((c) => c.dataset.fcStyle)).toEqual(CELL_STYLES.map((s) => s.id));
    handle.detach();
  });

  it('clicking a non-normal chip applies the style format to the active range', () => {
    setRange(store, 2, 3, 2, 3);
    const handle = attachCellStylesGallery({ host, store });
    handle.open();
    const goodChip = chips().find((c) => c.dataset.fcStyle === 'good');
    goodChip?.click();
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 3 }));
    // The 'good' style format has both color and fill set.
    expect(fmt?.color).toBe('#006100');
    expect(fmt?.fill).toBe('#c6efce');
    expect(overlay()?.hidden).toBe(true);
    handle.detach();
  });

  it('clicking the overlay backdrop (overlay element itself) closes the gallery', () => {
    const handle = attachCellStylesGallery({ host, store });
    handle.open();
    const o = overlay();
    o?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    expect(o?.hidden).toBe(true);
    handle.detach();
  });

  it('Escape closes the gallery', () => {
    const handle = attachCellStylesGallery({ host, store });
    handle.open();
    const o = overlay();
    o?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(o?.hidden).toBe(true);
    handle.detach();
  });

  it("clicking 'normal' clears every format field on the active range", () => {
    // Pre-stamp a format that should be wiped.
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true, color: '#ff0000' });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bold).toBe(
      true,
    );
    const handle = attachCellStylesGallery({ host, store });
    handle.open();
    const normalChip = chips().find((c) => c.dataset.fcStyle === 'normal');
    normalChip?.click();
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBeUndefined();
    expect(fmt?.color).toBeUndefined();
    handle.detach();
  });

  it('detach removes the overlay from the DOM', () => {
    const handle = attachCellStylesGallery({ host, store });
    handle.detach();
    expect(overlay()).toBeNull();
  });
});
