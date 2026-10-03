import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { setActive, setRange } from './fixtures.js';

describe('attachFormatDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setActive(store, 0, 0);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('border presets toggle all sides', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    // The border tab buttons appear after the color reset button. Locate by label.
    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const presetOutline = buttons.find((b) => b.textContent === '外枠') as HTMLButtonElement;
    const presetNone = buttons.find((b) => b.textContent === 'なし') as HTMLButtonElement;
    const presetAll = buttons.find((b) => b.textContent === '格子') as HTMLButtonElement;
    const presetRow = presetOutline.parentElement;

    expect(presetRow?.classList.contains('fc-fmtdlg__border-presets')).toBe(true);
    expect(presetNone.classList.contains('fc-fmtdlg__border-preset--none')).toBe(true);
    expect(presetOutline.classList.contains('fc-fmtdlg__border-preset--outline')).toBe(true);
    expect(presetAll.classList.contains('fc-fmtdlg__border-preset--inside')).toBe(true);

    presetOutline.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    let okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    let borders = store
      .getState()
      .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    const thinSide = { style: 'thin' };
    expect(borders).toEqual({
      top: thinSide,
      right: thinSide,
      bottom: thinSide,
      left: thinSide,
    });

    handle.open();
    presetNone.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    borders = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    expect(borders).toEqual({
      top: false,
      right: false,
      bottom: false,
      left: false,
      diagonalDown: false,
      diagonalUp: false,
    });

    handle.open();
    presetAll.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    borders = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    expect(borders).toEqual({
      top: thinSide,
      right: thinSide,
      bottom: thinSide,
      left: thinSide,
      diagonalDown: false,
      diagonalUp: false,
    });

    handle.detach();
  });

  it('border Outline preset applies only the perimeter of a multi-cell selection', () => {
    setRange(store, 0, 0, 1, 1);
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const presetOutline = buttons.find((b) => b.textContent === '外枠') as HTMLButtonElement;
    presetOutline.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats;
    const topLeft = fmt.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    const topRight = fmt.get(addrKey({ sheet: 0, row: 0, col: 1 }))?.borders;
    const bottomLeft = fmt.get(addrKey({ sheet: 0, row: 1, col: 0 }))?.borders;
    const bottomRight = fmt.get(addrKey({ sheet: 0, row: 1, col: 1 }))?.borders;
    const thinSide = { style: 'thin' };
    expect(topLeft).toMatchObject({ top: thinSide, left: thinSide });
    expect(topLeft?.right).toBeUndefined();
    expect(topLeft?.bottom).toBeUndefined();
    expect(topRight).toMatchObject({ top: thinSide, right: thinSide });
    expect(bottomLeft).toMatchObject({ bottom: thinSide, left: thinSide });
    expect(bottomRight).toMatchObject({ bottom: thinSide, right: thinSide });
    handle.detach();
  });

  it('border Outline preset does not materialize huge selections', () => {
    setRange(store, 0, 0, 100_000, 0);
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const presetOutline = buttons.find((b) => b.textContent === '外枠') as HTMLButtonElement;
    presetOutline.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().format.formats.size).toBe(0);
    handle.detach();
  });

  it('visual line style picker sets the active border style', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const thick = document.querySelector<HTMLButtonElement>(
      'button[data-border-style="thick"]',
    ) as HTMLButtonElement;
    const top = document.querySelector<HTMLButtonElement>(
      'button[data-border-side="top"]',
    ) as HTMLButtonElement;
    thick.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    top.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(thick.getAttribute('aria-pressed')).toBe('true');

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const borders = store
      .getState()
      .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    expect(borders?.top).toEqual({ style: 'thick' });
    handle.detach();
  });

  it('keeps existing per-side styles and colors when the active border pen changes', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        borders: {
          top: { style: 'thick', color: '#112233' },
          bottom: { style: 'hair', color: '#445566' },
        },
      },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();

    document
      .querySelector<HTMLButtonElement>('button[data-border-style="double"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    document
      .querySelector<HTMLButtonElement>('[data-swatches="border"] button[data-color="#c00000"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    document
      .querySelector<HTMLButtonElement>('button[data-border-side="left"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders,
    ).toMatchObject({
      top: { style: 'thick', color: '#112233' },
      bottom: { style: 'hair', color: '#445566' },
      left: { style: 'double', color: '#c00000' },
    });
    handle.detach();
  });

  it('reopens all OOXML border styles without collapsing the selected pen', () => {
    const styles = [
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
      'hair',
      'mediumDashed',
      'dashDot',
      'mediumDashDot',
      'dashDotDot',
      'mediumDashDotDot',
      'slantDashDot',
    ] as const;

    for (const style of styles) {
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 0 },
        {
          borders: { top: { style } },
        },
      );
      const handle = attachFormatDialog({ host, store });
      handle.open();
      expect(
        document.querySelector<HTMLSelectElement>(
          '.fc-fmtdlg__panel-tab[data-fc-tab="border"] select',
        )?.value,
      ).toBe(style);
      handle.detach();
    }
  });

  it('visual border preview buttons toggle sides', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const top = document.querySelector<HTMLButtonElement>(
      'button[data-border-side="top"]',
    ) as HTMLButtonElement;
    const left = document.querySelector<HTMLButtonElement>(
      'button[data-border-side="left"]',
    ) as HTMLButtonElement;
    top.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    left.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(top.getAttribute('aria-pressed')).toBe('true');
    expect(left.getAttribute('aria-pressed')).toBe('true');

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const borders = store
      .getState()
      .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    expect(borders?.top).toEqual({ style: 'thin' });
    expect(borders?.left).toEqual({ style: 'thin' });

    handle.detach();
  });

  it('border color swatches set the active line color', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const borderSwatch = document.querySelector<HTMLButtonElement>(
      '[data-swatches="border"] button[data-color="#c00000"]',
    ) as HTMLButtonElement;
    const bottom = document.querySelector<HTMLButtonElement>(
      'button[data-border-side="bottom"]',
    ) as HTMLButtonElement;
    borderSwatch.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    bottom.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const borders = store
      .getState()
      .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    expect(borders?.bottom).toEqual({ style: 'thin', color: '#c00000' });

    handle.detach();
  });
});
