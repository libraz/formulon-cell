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

  it('fill color input + reset', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fillInput = document.querySelector<HTMLInputElement>(
      'input[data-fc-color="fill"]',
    ) as HTMLInputElement;
    fillInput.value = '#123456';
    fillInput.dispatchEvent(new Event('input', { bubbles: true }));

    let okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill).toBe(
      '#123456',
    );

    handle.open();
    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const fillReset = buttons.find((b) => b.textContent === '塗りつぶしなし') as HTMLButtonElement;
    fillReset.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill,
    ).toBeUndefined();

    handle.detach();
  });

  it('Fill tab persists pattern style and pattern color', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="fill"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const pattern = document.querySelector<HTMLSelectElement>(
      'select[data-fc-select="fillPattern"]',
    );
    const patternColor = document.querySelector<HTMLInputElement>(
      'input[data-fc-color="fillPattern"]',
    );
    const sample = document.querySelector<HTMLDivElement>('.fc-fmtdlg__fill-sample-box');
    if (!pattern || !patternColor) throw new Error('fill pattern controls missing');
    if (!sample) throw new Error('fill sample missing');
    const fillInput = document.querySelector<HTMLInputElement>(
      'input[data-fc-color="fill"]',
    ) as HTMLInputElement;
    fillInput.value = '#123456';
    fillInput.dispatchEvent(new Event('input', { bubbles: true }));
    pattern.value = 'darkTrellis';
    pattern.dispatchEvent(new Event('change', { bubbles: true }));
    patternColor.value = '#336699';
    patternColor.dispatchEvent(new Event('input', { bubbles: true }));

    expect(sample.style.backgroundColor).toBe('#123456');
    expect(sample.style.backgroundImage).toContain('repeating-linear-gradient');
    expect(sample.style.backgroundImage).toContain('#336699');

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.fillPattern).toBe('darkTrellis');
    expect(fmt?.fillPatternColor).toBe('#336699');
    handle.detach();
  });

  it('uses the 6 by 3 pattern gallery to update the fill pattern', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('fill');
    const gallery = document.querySelector<HTMLDivElement>('.fc-fmtdlg__fill-pattern-gallery');
    const darkGrid = document.querySelector<HTMLButtonElement>(
      'button[data-fc-fill-pattern="darkGrid"]',
    );
    expect(gallery?.querySelectorAll('button[data-fc-fill-pattern]')).toHaveLength(18);
    darkGrid?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(darkGrid?.getAttribute('aria-pressed')).toBe('true');
    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fillPattern,
    ).toBe('darkGrid');
    handle.detach();
  });

  it('renders Format Cells select controls through the shared dialog select contract', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('align');

    const hAlign = document.querySelector<HTMLSelectElement>('select[data-fc-select="align"]');
    const vAlign = document.querySelector<HTMLSelectElement>('select[data-fc-select="vAlign"]');
    const textDirection = document.querySelector<HTMLSelectElement>(
      'select[data-fc-select="textDirection"]',
    );
    expect(hAlign?.getAttribute('aria-label')).toBe('横位置');
    expect(Array.from(hAlign?.options ?? []).map((option) => option.value)).toEqual([
      'default',
      'left',
      'center',
      'right',
      'fill',
      'justify',
      'centerContinuous',
      'distributed',
    ]);
    expect(Array.from(vAlign?.options ?? []).map((option) => option.value)).toEqual([
      'default',
      'top',
      'middle',
      'bottom',
      'justify',
      'distributed',
    ]);
    expect(Array.from(textDirection?.options ?? []).map((option) => option.value)).toEqual([
      'context',
      'ltr',
      'rtl',
    ]);

    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="fill"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const fillPattern = document.querySelector<HTMLSelectElement>(
      'select[data-fc-select="fillPattern"]',
    );
    expect(fillPattern?.getAttribute('aria-label')).toBe('パターンの種類');
    expect(Array.from(fillPattern?.options ?? []).map((option) => option.value)).toEqual(
      expect.arrayContaining(['darkUp', 'lightTrellis', 'gray0625']),
    );
    expect(fillPattern?.options).toHaveLength(18);

    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="border"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const borderStyle = document.querySelector<HTMLSelectElement>(
      '.fc-fmtdlg__panel-tab[data-fc-tab="border"] select',
    );
    expect(borderStyle?.getAttribute('aria-label')).toBe('スタイル');
    expect(Array.from(borderStyle?.options ?? []).map((option) => option.value)).toEqual([
      'thin',
      'hair',
      'dotted',
      'dashed',
      'dashDot',
      'dashDotDot',
      'mediumDashed',
      'mediumDashDot',
      'mediumDashDotDot',
      'medium',
      'double',
      'thick',
      'slantDashDot',
    ]);

    handle.detach();
  });

  it('font and fill swatches apply colors', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const fontSwatch = document.querySelector<HTMLButtonElement>(
      '[data-swatches="font"] button[data-color="#0070c0"]',
    ) as HTMLButtonElement;
    const fillSwatch = document.querySelector<HTMLButtonElement>(
      '[data-swatches="fill"] button[data-color="#ffff00"]',
    ) as HTMLButtonElement;
    fontSwatch.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    fillSwatch.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.color).toBe('#0070c0');
    expect(fmt?.fill).toBe('#ffff00');

    handle.detach();
  });

  it('hydrates fill picker default for non-hex existing fill', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fill: 'red' });
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const fillInput = document.querySelector<HTMLInputElement>('input[data-fc-color="fill"]');
    expect(fillInput?.value).toBe('#ffffff');
    handle.detach();
  });
});
