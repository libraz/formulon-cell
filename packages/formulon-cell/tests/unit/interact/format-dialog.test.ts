import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../src/interact/format-dialog.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const setActive = (store: SpreadsheetStore, row: number, col: number): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row, col },
      anchor: { sheet: 0, row, col },
      range: { sheet: 0, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

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
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const flushRaf = (): Promise<void> =>
  new Promise<void>((resolve) => {
    requestAnimationFrame(() => resolve());
  });

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

  it('mounts a hidden overlay on attach', () => {
    const handle = attachFormatDialog({ host, store });
    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay).not.toBeNull();
    expect(overlay?.hidden).toBe(true);
    handle.detach();
  });

  it('open() reveals the overlay and focuses the active tab button', async () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay?.hidden).toBe(false);

    await flushRaf();
    const numberTab = host.querySelector<HTMLButtonElement>('button[data-fc-tab="number"]');
    expect(document.activeElement).toBe(numberTab);
    handle.detach();
  });

  it('close() hides the overlay and refocuses host', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    handle.close();
    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay?.hidden).toBe(true);
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('detach() removes the overlay from DOM', () => {
    const handle = attachFormatDialog({ host, store });
    expect(host.querySelector('.fc-fmtdlg')).not.toBeNull();
    handle.detach();
    expect(host.querySelector('.fc-fmtdlg')).toBeNull();
  });

  it('hydrates draft from active cell format on open', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 4 },
        align: 'right',
        bold: true,
        italic: true,
        underline: true,
        strike: true,
        fontFamily: 'Georgia',
        fontSize: 14,
        color: '#ff0000',
        fill: '#00ff00',
        borders: { top: true, right: true, bottom: false, left: false },
      },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const decimalsInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('4');

    const boldInput = host.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
    expect(boldInput?.checked).toBe(true);

    const familyInput = host.querySelector<HTMLInputElement>('input[data-fc-input="family"]');
    expect(familyInput?.value).toBe('Georgia');

    const sizeInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="8"][max="72"]',
    );
    expect(sizeInput?.value).toBe('14');

    const rightAlign = host.querySelector<HTMLInputElement>('input[type="radio"][value="right"]');
    expect(rightAlign?.checked).toBe(true);

    handle.detach();
  });

  it('hydrates draft from currency numFmt with symbol', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { numFmt: { kind: 'currency', decimals: 3, symbol: '€' } },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const decimalsInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('3');
    const symbolSelect = host.querySelector<HTMLSelectElement>('select');
    expect(symbolSelect?.value).toBe('€');
    handle.detach();
  });

  it('hydrates draft from percent numFmt', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { numFmt: { kind: 'percent', decimals: 1 } },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const decimalsInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('1');
    handle.detach();
  });

  it('falls back to general/defaults when active cell has no format', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const decimalsInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    );
    expect(decimalsInput?.value).toBe('2');
    const symbolSelect = host.querySelector<HTMLSelectElement>('select');
    expect(symbolSelect?.value).toBe('$');
    handle.detach();
  });

  it('switching tabs toggles aria-selected and panel visibility', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fontTab = host.querySelector<HTMLButtonElement>('button[data-fc-tab="font"]');
    fontTab?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(fontTab?.getAttribute('aria-selected')).toBe('true');
    const numberTab = host.querySelector<HTMLButtonElement>('button[data-fc-tab="number"]');
    expect(numberTab?.getAttribute('aria-selected')).toBe('false');

    const fontPanel = host.querySelector<HTMLDivElement>('div[data-fc-tab="font"]');
    const numberPanel = host.querySelector<HTMLDivElement>('div[data-fc-tab="number"]');
    expect(fontPanel?.hidden).toBe(false);
    expect(numberPanel?.hidden).toBe(true);

    handle.detach();
  });

  it('clicking tab strip outside button does nothing', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const tabsStrip = host.querySelector<HTMLElement>('.fc-fmtdlg__tabs');
    tabsStrip?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const numberTab = host.querySelector<HTMLButtonElement>('button[data-fc-tab="number"]');
    expect(numberTab?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('clicking number category updates draft and visibility', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const decimalsRow = host.querySelector<HTMLLabelElement>(
      '.fc-fmtdlg__cat-controls .fc-fmtdlg__row',
    );
    expect(decimalsRow?.hidden).toBe(true); // general → hidden

    const fixedBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(fixedBtn?.getAttribute('aria-selected')).toBe('true');
    expect(decimalsRow?.hidden).toBe(false);

    const symbolRow = host.querySelectorAll<HTMLLabelElement>(
      '.fc-fmtdlg__cat-controls .fc-fmtdlg__row',
    )[1];
    expect(symbolRow?.hidden).toBe(true);

    const currencyBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="currency"]');
    currencyBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(symbolRow?.hidden).toBe(false);

    const percentBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="percent"]');
    percentBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(symbolRow?.hidden).toBe(true);
    expect(decimalsRow?.hidden).toBe(false);

    handle.detach();
  });

  it('clicking on cat list outside button is a no-op', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const catList = host.querySelector<HTMLElement>('.fc-fmtdlg__cat');
    catList?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const generalBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="general"]');
    expect(generalBtn?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('decimals input clamps to [0, 10]', () => {
    const history = new History();
    const handle = attachFormatDialog({ host, store, history });
    handle.open();
    const fixedBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const decimalsInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    ) as HTMLInputElement;
    decimalsInput.value = '99';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'fixed', decimals: 10 });
    handle.detach();
  });

  it('decimals input ignores non-numeric values', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const fixedBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const decimalsInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    ) as HTMLInputElement;
    decimalsInput.value = 'abc';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
    handle.detach();
  });

  it('symbol select updates currency symbol', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const currencyBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="currency"]');
    currencyBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const symbolSelect = host.querySelector<HTMLSelectElement>('select') as HTMLSelectElement;
    symbolSelect.value = '¥';
    symbolSelect.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'currency', decimals: 2, symbol: '¥' });
    handle.detach();
  });

  it('alignment radios update draft (default → undefined)', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const center = host.querySelector<HTMLInputElement>(
      'input[type="radio"][value="center"]',
    ) as HTMLInputElement;
    center.checked = true;
    center.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.align).toBe(
      'center',
    );
    handle.detach();
  });

  it('alignment "default" radio clears align', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { align: 'right' });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const dflt = host.querySelector<HTMLInputElement>(
      'input[type="radio"][value="default"]',
    ) as HTMLInputElement;
    dflt.checked = true;
    dflt.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.align,
    ).toBeUndefined();
    handle.detach();
  });

  it('unchecked alignment radio change is ignored', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const left = host.querySelector<HTMLInputElement>(
      'input[type="radio"][value="left"]',
    ) as HTMLInputElement;
    left.checked = false;
    left.dispatchEvent(new Event('change', { bubbles: true }));
    handle.detach();
  });

  it('font style checkboxes wire to draft', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const bold = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    const italic = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="italic"]',
    ) as HTMLInputElement;
    const underline = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="underline"]',
    ) as HTMLInputElement;
    const strike = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="strike"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    italic.checked = true;
    italic.dispatchEvent(new Event('change', { bubbles: true }));
    underline.checked = true;
    underline.dispatchEvent(new Event('change', { bubbles: true }));
    strike.checked = true;
    strike.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(true);
    expect(fmt?.italic).toBe(true);
    expect(fmt?.underline).toBe(true);
    expect(fmt?.strike).toBe(true);
    handle.detach();
  });

  it('font family input sets draft and applies on OK', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const familyInput = host.querySelector<HTMLInputElement>(
      'input[data-fc-input="family"]',
    ) as HTMLInputElement;
    familyInput.value = 'Helvetica';
    familyInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontFamily,
    ).toBe('Helvetica');
    handle.detach();
  });

  it('empty font family is converted to undefined on OK', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontFamily: 'Arial' });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const familyInput = host.querySelector<HTMLInputElement>(
      'input[data-fc-input="family"]',
    ) as HTMLInputElement;
    familyInput.value = '';
    familyInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontFamily,
    ).toBeUndefined();
    handle.detach();
  });

  it('font size input clamps to [8, 72]', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const sizeInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="8"][max="72"]',
    ) as HTMLInputElement;
    sizeInput.value = '500';
    sizeInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontSize,
    ).toBe(72);
    handle.detach();
  });

  it('empty font size becomes undefined on OK', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontSize: 18 });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const sizeInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="8"][max="72"]',
    ) as HTMLInputElement;
    sizeInput.value = '';
    sizeInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontSize,
    ).toBeUndefined();
    handle.detach();
  });

  it('non-numeric font size leaves draft unchanged', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontSize: 12 });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const sizeInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="8"][max="72"]',
    ) as HTMLInputElement;
    // Bypass happy-dom number-input value coercion by overriding the getter
    // so the handler observes a non-empty NaN-parseable string.
    Object.defineProperty(sizeInput, 'value', {
      configurable: true,
      get: () => 'abc',
      set: () => {},
    });
    sizeInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontSize,
    ).toBe(12);
    handle.detach();
  });

  it('color input sets draft, color reset clears it', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const colorInputs = host.querySelectorAll<HTMLInputElement>('input[type="color"]');
    const colorInput = colorInputs[0] as HTMLInputElement;
    colorInput.value = '#abcdef';
    colorInput.dispatchEvent(new Event('input', { bubbles: true }));

    let okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.color).toBe(
      '#abcdef',
    );

    handle.open();
    // The first non-primary button in font panel is "reset to default" for color
    const colorReset = host.querySelectorAll<HTMLButtonElement>(
      '.fc-fmtdlg__btn',
    )[0] as HTMLButtonElement;
    colorReset.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.color,
    ).toBeUndefined();

    handle.detach();
  });

  it('hydrates color picker default for non-hex existing color', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { color: 'rebeccapurple' });
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const colorInputs = host.querySelectorAll<HTMLInputElement>('input[type="color"]');
    expect(colorInputs[0]?.value).toBe('#000000');
    handle.detach();
  });

  it('border presets toggle all sides', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    // The border tab buttons appear after the color reset button. Locate by label.
    const buttons = Array.from(host.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const presetOutline = buttons.find((b) => b.textContent === '外枠') as HTMLButtonElement;
    const presetNone = buttons.find((b) => b.textContent === 'なし') as HTMLButtonElement;
    const presetAll = buttons.find((b) => b.textContent === '格子') as HTMLButtonElement;

    presetOutline.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    let okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
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
      diagonalDown: false,
      diagonalUp: false,
    });

    handle.open();
    presetNone.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
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
    okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
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

  it('individual border checkboxes flip sides', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const top = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="borderTop"]',
    ) as HTMLInputElement;
    const bottom = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="borderBottom"]',
    ) as HTMLInputElement;
    const left = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="borderLeft"]',
    ) as HTMLInputElement;
    const right = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="borderRight"]',
    ) as HTMLInputElement;

    top.checked = true;
    top.dispatchEvent(new Event('change', { bubbles: true }));
    bottom.checked = true;
    bottom.dispatchEvent(new Event('change', { bubbles: true }));
    left.checked = true;
    left.dispatchEvent(new Event('change', { bubbles: true }));
    right.checked = true;
    right.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const borders = store
      .getState()
      .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.borders;
    const thinSide = { style: 'thin' };
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

  it('fill color input + reset', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fillInput = host.querySelector<HTMLInputElement>(
      'input[data-fc-color="fill"]',
    ) as HTMLInputElement;
    fillInput.value = '#123456';
    fillInput.dispatchEvent(new Event('input', { bubbles: true }));

    let okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill).toBe(
      '#123456',
    );

    handle.open();
    const buttons = Array.from(host.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const fillReset = buttons.find((b) => b.textContent === '塗りつぶしなし') as HTMLButtonElement;
    fillReset.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill,
    ).toBeUndefined();

    handle.detach();
  });

  it('hydrates fill picker default for non-hex existing fill', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fill: 'red' });
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const fillInput = host.querySelector<HTMLInputElement>('input[data-fc-color="fill"]');
    expect(fillInput?.value).toBe('#ffffff');
    handle.detach();
  });

  it('Cancel button closes the dialog without writing', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const buttons = Array.from(host.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const cancelBtn = buttons.find((b) => b.textContent === 'キャンセル') as HTMLButtonElement;
    cancelBtn.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg');
    expect(overlay?.hidden).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toBeUndefined();
    handle.detach();
  });

  it('backdrop click closes the dialog; panel click does not', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const panel = host.querySelector<HTMLElement>('.fc-fmtdlg__panel') as HTMLElement;

    panel.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(overlay.hidden).toBe(false);

    overlay.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(overlay.hidden).toBe(true);

    handle.detach();
  });

  it('Escape closes the dialog and stops propagation', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const hostKey = vi.fn();
    host.addEventListener('keydown', hostKey);

    const ev = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true });
    overlay.dispatchEvent(ev);

    expect(overlay.hidden).toBe(true);
    expect(hostKey).not.toHaveBeenCalled();
    handle.detach();
  });

  it('Enter on input applies and closes', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fixedBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const decimalsInput = host.querySelector<HTMLInputElement>(
      'input[type="number"][min="0"][max="10"]',
    ) as HTMLInputElement;
    decimalsInput.value = '5';
    decimalsInput.dispatchEvent(new Event('input', { bubbles: true }));

    const ev = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true });
    Object.defineProperty(ev, 'target', { value: decimalsInput });
    overlay.dispatchEvent(ev);

    expect(overlay.hidden).toBe(true);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.numFmt).toEqual({ kind: 'fixed', decimals: 5 });

    handle.detach();
  });

  it('Enter on a button does not apply (lets button click handle it)', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const overlay = host.querySelector<HTMLElement>('.fc-fmtdlg') as HTMLElement;
    const someBtn = host.querySelector<HTMLButtonElement>(
      'button[data-fc-tab="font"]',
    ) as HTMLButtonElement;

    const ev = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true });
    Object.defineProperty(ev, 'target', { value: someBtn });
    overlay.dispatchEvent(ev);

    expect(overlay.hidden).toBe(false);
    handle.detach();
  });

  it('history records the format change as a single undo entry', () => {
    const history = new History();
    const handle = attachFormatDialog({ host, store, history });
    handle.open();

    const center = host.querySelector<HTMLInputElement>(
      'input[type="radio"][value="center"]',
    ) as HTMLInputElement;
    center.checked = true;
    center.dispatchEvent(new Event('change', { bubbles: true }));

    const bold = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    let fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.align).toBe('center');
    expect(fmt?.bold).toBe(true);

    expect(history.undo()).toBe(true);
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.align).toBeUndefined();
    expect(fmt?.bold).toBeFalsy();

    expect(history.redo()).toBe(true);
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.align).toBe('center');
    expect(fmt?.bold).toBe(true);

    handle.detach();
  });

  it('applies patch over the entire selection range', () => {
    setRange(store, 0, 0, 1, 1);
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const bold = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = host.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const formats = store.getState().format.formats;
    expect(formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bold).toBe(true);
    expect(formats.get(addrKey({ sheet: 0, row: 0, col: 1 }))?.bold).toBe(true);
    expect(formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))?.bold).toBe(true);
    expect(formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))?.bold).toBe(true);
    handle.detach();
  });

  it('preview reflects current draft and includes formatted number', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const preview = host.querySelector<HTMLElement>('.fc-fmtdlg__preview') as HTMLElement;

    const fixedBtn = host.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
    fixedBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(preview.textContent).toMatch(/12,?345\.00/);

    const bold = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    const italic = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="italic"]',
    ) as HTMLInputElement;
    const underline = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="underline"]',
    ) as HTMLInputElement;
    const strike = host.querySelector<HTMLInputElement>(
      'input[data-fc-check="strike"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.fontWeight).toBe('bold');

    italic.checked = true;
    italic.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.fontStyle).toBe('italic');

    underline.checked = true;
    underline.dispatchEvent(new Event('change', { bubbles: true }));
    strike.checked = true;
    strike.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.textDecoration).toMatch(/underline/);
    expect(preview.style.textDecoration).toMatch(/line-through/);

    const right = host.querySelector<HTMLInputElement>(
      'input[type="radio"][value="right"]',
    ) as HTMLInputElement;
    right.checked = true;
    right.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.textAlign).toBe('right');

    handle.detach();
  });

  it('detach() removes wired listeners (clicking after detach is inert)', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    handle.detach();
    expect(host.querySelector('.fc-fmtdlg')).toBeNull();
  });
});
