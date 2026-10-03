import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { setActive } from './fixtures.js';

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

  it('font style controls wire the selected underline style to draft', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const bold = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    const italic = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="italic"]',
    ) as HTMLInputElement;
    const underline = document.querySelector<HTMLSelectElement>(
      'select[data-fc-input="underline"]',
    ) as HTMLSelectElement;
    const strike = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="strike"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    italic.checked = true;
    italic.dispatchEvent(new Event('change', { bubbles: true }));
    underline.value = 'doubleAccounting';
    underline.dispatchEvent(new Event('change', { bubbles: true }));
    strike.checked = true;
    strike.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(true);
    expect(fmt?.italic).toBe(true);
    expect(fmt?.underline).toBe('doubleAccounting');
    expect(fmt?.strike).toBe(true);
    handle.detach();
  });

  it('font vertical-alignment controls are mutually exclusive and write the selected value', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const superscript = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="superscript"]',
    ) as HTMLInputElement;
    const subscript = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="subscript"]',
    ) as HTMLInputElement;
    superscript.checked = true;
    superscript.dispatchEvent(new Event('change', { bubbles: true }));
    expect(subscript.checked).toBe(false);
    subscript.checked = true;
    subscript.dispatchEvent(new Event('change', { bubbles: true }));
    expect(superscript.checked).toBe(false);

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontVertAlign,
    ).toBe('subscript');
    handle.detach();
  });

  it('Font tab style list applies Regular, Italic, Bold, and Bold Italic choices', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="font"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const normal = document.querySelector<HTMLInputElement>('input[data-fc-check="normalFont"]');
    const boldItalic = document.querySelector<HTMLButtonElement>(
      'button[data-fc-font-style="boldItalic"]',
    );
    if (!normal || !boldItalic) throw new Error('font style controls missing');
    expect(normal.checked).toBe(true);

    boldItalic.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(normal.checked).toBe(false);
    expect(boldItalic.getAttribute('aria-selected')).toBe('true');

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(true);
    expect(fmt?.italic).toBe(true);
    handle.detach();
  });

  it('Font tab preview follows font family, size, style, and color draft', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('font');

    const familyInput = document.querySelector<HTMLInputElement>('input[data-fc-input="family"]');
    const sizeInput = document.querySelector<HTMLInputElement>('input[type="number"][min="1"]');
    const colorInput = document.querySelector<HTMLInputElement>('input[data-fc-color="font"]');
    const boldItalic = document.querySelector<HTMLButtonElement>(
      'button[data-fc-font-style="boldItalic"]',
    );
    const preview = document.querySelector<HTMLDivElement>('.fc-fmtdlg__font-preview-box');
    if (!familyInput || !sizeInput || !colorInput || !boldItalic || !preview) {
      throw new Error('font preview controls missing');
    }

    familyInput.value = 'Georgia';
    familyInput.dispatchEvent(new Event('input', { bubbles: true }));
    sizeInput.value = '18';
    sizeInput.dispatchEvent(new Event('input', { bubbles: true }));
    colorInput.value = '#cc0000';
    colorInput.dispatchEvent(new Event('input', { bubbles: true }));
    boldItalic.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(preview.style.fontFamily).toBe('Georgia');
    expect(preview.style.fontSize).toBe('18px');
    expect(preview.style.color).toBe('#cc0000');
    expect(preview.style.fontWeight).toBe('bold');
    expect(preview.style.fontStyle).toBe('italic');

    handle.detach();
  });

  it('Normal font resets font family, size, color, and effects', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        italic: true,
        underline: true,
        strike: true,
        fontFamily: 'Georgia',
        fontSize: 18,
        color: '#ff0000',
      },
    );
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="font"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const normal = document.querySelector<HTMLInputElement>('input[data-fc-check="normalFont"]');
    if (!normal) throw new Error('normal font checkbox missing');
    normal.checked = true;
    normal.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(false);
    expect(fmt?.italic).toBe(false);
    expect(fmt?.underline).toBe(false);
    expect(fmt?.strike).toBe(false);
    expect(fmt?.fontFamily).toBeUndefined();
    expect(fmt?.fontSize).toBeUndefined();
    expect(fmt?.color).toBeUndefined();
    handle.detach();
  });

  it('font family input sets draft and applies on OK', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const familyInput = document.querySelector<HTMLInputElement>(
      'input[data-fc-input="family"]',
    ) as HTMLInputElement;
    familyInput.value = 'Helvetica';
    familyInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
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

    const familyInput = document.querySelector<HTMLInputElement>(
      'input[data-fc-input="family"]',
    ) as HTMLInputElement;
    familyInput.value = '';
    familyInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontFamily,
    ).toBeUndefined();
    handle.detach();
  });

  it('font size input clamps to [1, 409]', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const sizeInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="1"][max="409"]',
    ) as HTMLInputElement;
    sizeInput.value = '500';
    sizeInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontSize,
    ).toBe(409);
    handle.detach();
  });

  it('accepts the Excel integer font-size bounds', () => {
    const handle = attachFormatDialog({ host, store });
    for (const size of [1, 7, 100, 409]) {
      handle.open('font');
      const sizeInput = document.querySelector<HTMLInputElement>(
        'input[type="number"][min="1"][max="409"]',
      );
      if (!sizeInput) throw new Error('font size input missing');
      sizeInput.value = String(size);
      sizeInput.dispatchEvent(new Event('input', { bubbles: true }));
      document
        .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
        ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontSize,
      ).toBe(size);
    }
    handle.detach();
  });

  it('empty font size becomes undefined on OK', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontSize: 18 });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const sizeInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="1"][max="409"]',
    ) as HTMLInputElement;
    sizeInput.value = '';
    sizeInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
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

    const sizeInput = document.querySelector<HTMLInputElement>(
      'input[type="number"][min="1"][max="409"]',
    ) as HTMLInputElement;
    // Bypass happy-dom number-input value coercion by overriding the getter
    // so the handler observes a non-empty NaN-parseable string.
    Object.defineProperty(sizeInput, 'value', {
      configurable: true,
      get: () => 'abc',
      set: () => {},
    });
    sizeInput.dispatchEvent(new Event('input', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fontSize,
    ).toBe(12);
    handle.detach();
  });

  it('color input sets draft, color reset clears it', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const colorInputs = document.querySelectorAll<HTMLInputElement>('input[type="color"]');
    const colorInput = colorInputs[0] as HTMLInputElement;
    colorInput.value = '#abcdef';
    colorInput.dispatchEvent(new Event('input', { bubbles: true }));

    let okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.color).toBe(
      '#abcdef',
    );

    handle.open();
    // The first non-primary button in font panel is "reset to default" for color
    const colorReset = document.querySelectorAll<HTMLButtonElement>(
      '.fc-fmtdlg__btn',
    )[0] as HTMLButtonElement;
    colorReset.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
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
    const colorInputs = document.querySelectorAll<HTMLInputElement>('input[type="color"]');
    expect(colorInputs[0]?.value).toBe('#000000');
    handle.detach();
  });
});
