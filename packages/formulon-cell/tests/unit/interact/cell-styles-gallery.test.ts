import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { CELL_STYLE_GROUPS, CELL_STYLES } from '../../../src/commands/cell-styles.js';
import { addrKey } from '../../../src/engine/address.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachCellStylesGallery } from '../../../src/interact/cell-styles-gallery.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const overlay = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-stylegallery');
const chips = (): HTMLButtonElement[] =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-stylegallery__chip'));
const headings = (): string[] =>
  Array.from(document.querySelectorAll<HTMLElement>('.fc-stylegallery__heading')).map((h) =>
    (h.textContent ?? '').trim(),
  );

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

  it('renders Excel-style groups with one chip per CELL_STYLES entry', () => {
    const handle = attachCellStylesGallery({ host, store, strings: en });
    handle.open();
    expect(
      document.querySelector<HTMLElement>('.fc-stylegallery__grid')?.getAttribute('role'),
    ).toBe('toolbar');
    expect(headings()).toEqual([
      'Good, Bad and Neutral',
      'Data and Model',
      'Titles and Headings',
      'Themed Cell Styles',
      'Number Format',
    ]);
    const list = chips();
    const expectedOrder = CELL_STYLE_GROUPS.flatMap((group) => group.styleIds);
    expect(list.length).toBe(CELL_STYLES.length);
    expect(list.map((c) => c.dataset.fcStyle)).toEqual(expectedOrder);
    expect(list[0]?.tabIndex).toBe(0);
    expect(list[1]?.tabIndex).toBe(-1);
    handle.detach();
  });

  it('uses localized style labels and updates them on locale changes', () => {
    const handle = attachCellStylesGallery({ host, store, strings: ja });
    handle.open();
    expect(overlay()?.getAttribute('aria-label')).toBe('セル スタイル');
    expect(headings()).toContain('テーマのセル スタイル');
    expect(chips().find((c) => c.dataset.fcStyle === 'normal')?.textContent).toBe('標準');
    expect(chips().find((c) => c.dataset.fcStyle === 'checkCell')?.textContent).toBe(
      'チェック セル',
    );
    expect(chips().find((c) => c.dataset.fcStyle === 'accent1_20')?.textContent).toBe(
      '20% - アクセント1',
    );

    handle.setStrings(en);
    expect(overlay()?.getAttribute('aria-label')).toBe('Cell styles');
    expect(headings()).toContain('Themed Cell Styles');
    expect(chips().find((c) => c.dataset.fcStyle === 'normal')?.textContent).toBe('Normal');
    expect(chips().find((c) => c.dataset.fcStyle === 'checkCell')?.textContent).toBe(
      'Check Cell',
    );
    handle.detach();
  });

  it('moves chip focus with Excel-style arrow, Home, and End keys', () => {
    const handle = attachCellStylesGallery({ host, store });
    handle.open();
    const list = chips();
    if (!list[0] || !list[1]) throw new Error('expected style chips');
    list[0].focus();

    list[0].dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(document.activeElement).toBe(list[1]);
    expect(list[1].tabIndex).toBe(0);
    expect(list[0].tabIndex).toBe(-1);

    list[1].dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(document.activeElement).toBe(list[list.length - 1]);

    list[list.length - 1]?.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Home', bubbles: true }),
    );
    expect(document.activeElement).toBe(list[0]);
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
