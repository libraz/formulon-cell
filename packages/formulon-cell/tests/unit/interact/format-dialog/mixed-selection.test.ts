import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { flushRaf, setActive, setRange } from './fixtures.js';

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

  it('shows mixed union fields and applies only a touched alignment field', () => {
    const active = { sheet: 0, row: 0, col: 0 };
    const other = { sheet: 0, row: 0, col: 1 };
    const extra = { sheet: 0, row: 4, col: 4 };
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
        extraRanges: [{ sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 }],
      },
    }));
    mutators.setCellFormat(store, active, {
      bold: true,
      hyperlink: 'https://active.test',
      comment: 'active note',
    });
    mutators.setCellFormat(store, other, { bold: false, hyperlink: 'https://other.test' });
    mutators.setCellFormat(store, extra, { bold: false, comment: 'extra note' });
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('font');
      const bold = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
      expect(bold?.indeterminate).toBe(true);

      handle.open('align');
      const center = document.querySelector<HTMLInputElement>(
        'input[type="radio"][value="center"]',
      );
      if (!center) throw new Error('center alignment control missing');
      center.checked = true;
      center.dispatchEvent(new Event('change', { bubbles: true }));
      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

      expect(store.getState().format.formats.get(addrKey(active))).toMatchObject({
        align: 'center',
        bold: true,
        hyperlink: 'https://active.test',
        comment: 'active note',
      });
      expect(store.getState().format.formats.get(addrKey(other))).toMatchObject({
        align: 'center',
        bold: false,
        hyperlink: 'https://other.test',
      });
      expect(store.getState().format.formats.get(addrKey(extra))).toMatchObject({
        align: 'center',
        bold: false,
        comment: 'extra note',
      });
    } finally {
      handle.detach();
    }
  });

  it('unifies a mixed field when the user explicitly touches its active value', () => {
    setRange(store, 0, 0, 1, 0);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { bold: false });
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('font');
      const bold = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
      if (!bold) throw new Error('bold checkbox missing');
      expect(bold.indeterminate).toBe(true);
      bold.checked = true;
      bold.dispatchEvent(new Event('change', { bubbles: true }));
      expect(bold.indeterminate).toBe(false);
      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
      expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bold).toBe(
        true,
      );
      expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))?.bold).toBe(
        true,
      );
    } finally {
      handle.detach();
    }
  });

  it('lets the user choose None for a mixed underline', async () => {
    setRange(store, 0, 0, 0, 1);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { underline: 'single' });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { bold: true });
    const handle = attachFormatDialog({ host, store });
    try {
      await flushRaf();
      handle.open('font');
      await flushRaf();
      const select = document.querySelector<HTMLSelectElement>('select[data-fc-input="underline"]');
      if (!select) throw new Error('underline select missing');
      expect(select.value).not.toBe('');

      const button = select.parentElement?.querySelector<HTMLButtonElement>('.fc-select__button');
      button?.click();
      select.parentElement
        ?.querySelector<HTMLButtonElement>('.fc-select__option[data-value=""]')
        ?.click();
      expect(select.value).toBe('');

      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
      expect(
        store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.underline,
      ).toBeFalsy();
    } finally {
      handle.detach();
    }
  });

  it('clears mixed font style, family, and size galleries until each group is touched', () => {
    setRange(store, 0, 0, 0, 1);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        italic: false,
        fontFamily: 'Arial',
        fontSize: 7,
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        bold: false,
        italic: true,
        fontFamily: 'Georgia',
        fontSize: 409,
      },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('font');
      const styleItems = document.querySelectorAll<HTMLButtonElement>('[data-fc-font-style]');
      const familyItems = document.querySelectorAll<HTMLButtonElement>('[data-fc-font-family]');
      const sizeItems = document.querySelectorAll<HTMLButtonElement>('[data-fc-font-size]');
      expect([...styleItems].every((item) => item.getAttribute('aria-selected') === 'false')).toBe(
        true,
      );
      expect([...familyItems].every((item) => item.getAttribute('aria-selected') === 'false')).toBe(
        true,
      );
      expect([...sizeItems].every((item) => item.getAttribute('aria-selected') === 'false')).toBe(
        true,
      );
      expect(
        document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]')?.indeterminate,
      ).toBe(true);
      expect(
        document.querySelector<HTMLInputElement>('input[data-fc-check="italic"]')?.indeterminate,
      ).toBe(true);
      expect(
        document.querySelector<HTMLInputElement>('input[data-fc-check="normalFont"]')
          ?.indeterminate,
      ).toBe(true);

      document.querySelector<HTMLButtonElement>('[data-fc-font-style="boldItalic"]')?.click();
      document.querySelector<HTMLButtonElement>('[data-fc-font-family="Arial"]')?.click();
      document.querySelector<HTMLButtonElement>('[data-fc-font-size="18"]')?.click();
      expect(
        document
          .querySelector<HTMLButtonElement>('[data-fc-font-style="boldItalic"]')
          ?.getAttribute('aria-selected'),
      ).toBe('true');
      expect(
        document
          .querySelector<HTMLButtonElement>('[data-fc-font-family="Arial"]')
          ?.getAttribute('aria-selected'),
      ).toBe('true');
      expect(
        document
          .querySelector<HTMLButtonElement>('[data-fc-font-size="18"]')
          ?.getAttribute('aria-selected'),
      ).toBe('true');
      expect(
        document.querySelector<HTMLInputElement>('input[data-fc-input="family"]')?.dataset.fcMixed,
      ).toBe(undefined);
      expect(
        document.querySelector<HTMLInputElement>('input[type="number"][min="1"][max="409"]')
          ?.dataset.fcMixed,
      ).toBe(undefined);

      handle.close();
      setActive(store, 2, 2);
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 2, col: 2 },
        {
          bold: false,
          italic: false,
          fontFamily: 'Georgia',
          fontSize: 18,
        },
      );
      handle.open('font');
      expect(
        document
          .querySelector<HTMLButtonElement>('[data-fc-font-style="regular"]')
          ?.getAttribute('aria-selected'),
      ).toBe('true');
      expect(
        document
          .querySelector<HTMLButtonElement>('[data-fc-font-family="Georgia"]')
          ?.getAttribute('aria-selected'),
      ).toBe('true');
      expect(
        document
          .querySelector<HTMLButtonElement>('[data-fc-font-size="18"]')
          ?.getAttribute('aria-selected'),
      ).toBe('true');
      expect(
        [...document.querySelectorAll<HTMLButtonElement>('[data-fc-font-size]')].filter(
          (item) => item.getAttribute('aria-selected') === 'true',
        ),
      ).toHaveLength(1);
    } finally {
      handle.detach();
    }
  });

  it.each(['strike', 'color'] as const)(
    'clears Normal Font mixed state when the last mixed %s field is chosen',
    (field) => {
      setRange(store, 0, 0, 0, 1);
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 0 },
        field === 'strike' ? { strike: true } : { color: '#ff0000' },
      );
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 1 },
        field === 'strike' ? { strike: false } : { color: '#0000ff' },
      );
      const handle = attachFormatDialog({ host, store });
      try {
        handle.open('font');
        const normal = document.querySelector<HTMLInputElement>('[data-fc-check="normalFont"]');
        if (!normal) throw new Error('Normal Font control missing');
        expect(normal.indeterminate).toBe(true);
        expect(normal.dataset.fcMixed).toBe('true');
        const control = document.querySelector<HTMLElement>(
          field === 'strike'
            ? '[data-fc-check="strike"]'
            : '[data-swatches="font"] button[data-color="#0070c0"]',
        );
        if (!control) throw new Error('font field control missing');
        control.click();
        expect(normal.indeterminate).toBe(false);
        expect(normal.dataset.fcMixed).toBeUndefined();
        document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
        for (const col of [0, 1]) {
          expect(store.getState().format.formats.get(`0:0:${col}`)?.[field]).toBe(
            field === 'strike' ? false : '#0070c0',
          );
        }
      } finally {
        handle.detach();
      }
    },
  );

  it.each(['negative', 'pattern'] as const)(
    'ignores %s list padding without unifying mixed number formats',
    (list) => {
      setRange(store, 0, 0, 0, 1);
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 0 },
        { numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' } },
      );
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 1 },
        { numFmt: { kind: 'date', pattern: 'dd/mm/yyyy' } },
      );
      const before = store.getState().format.formats;
      const history = new History();
      const handle = attachFormatDialog({ host, store, history });
      try {
        handle.open('number');
        const item = document.querySelector<HTMLElement>(
          list === 'negative' ? '[data-fc-negative-style]' : '[data-fc-pattern]',
        );
        const container = item?.parentElement;
        if (!container) throw new Error('number format list missing');
        container.click();
        document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
        expect(store.getState().format.formats).toBe(before);
        expect(history.canUndo()).toBe(false);
      } finally {
        handle.detach();
      }
    },
  );

  it('projects mixed number, fill, and validation groups without selecting stale descendants', () => {
    setRange(store, 0, 0, 0, 1);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'fixed', decimals: 2, thousands: true, negativeStyle: 'red' },
        fill: '#ffeeaa',
        fillPattern: 'gray50',
        fillPatternColor: '#112233',
        validation: {
          kind: 'whole',
          op: '=',
          a: 1,
          allowBlank: false,
          showInputMessage: true,
          showErrorMessage: true,
        },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        numFmt: { kind: 'currency', decimals: 3, symbol: '€', negativeStyle: 'parens' },
        fill: '#ddeeff',
        fillPattern: 'darkGrid',
        fillPatternColor: '#445566',
        validation: { kind: 'list', source: ['A', 'B'] },
      },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('number');
      const numberPanel = document.querySelector<HTMLElement>(
        '.fc-fmtdlg__panel-tab[data-fc-tab="number"]',
      );
      if (!numberPanel) throw new Error('number panel missing');
      expect(
        numberPanel.querySelector<HTMLInputElement>('input[type="number"][min="0"][max="10"]')
          ?.value,
      ).toBe('');
      expect(
        [...numberPanel.querySelectorAll<HTMLButtonElement>('[data-fc-negative-style]')].every(
          (button) => button.getAttribute('aria-selected') === 'false',
        ),
      ).toBe(true);
      expect(
        [...numberPanel.querySelectorAll<HTMLButtonElement>('[data-fc-cat]')].every(
          (button) => button.getAttribute('aria-selected') === 'false',
        ),
      ).toBe(true);
      const symbol = [...numberPanel.querySelectorAll<HTMLSelectElement>('select')].find((select) =>
        [...select.options].some((option) => option.value === '€'),
      );
      expect(symbol?.value).toBe('');

      document.querySelector<HTMLButtonElement>('[data-fc-cat="fixed"]')?.click();
      expect(
        document
          .querySelector<HTMLButtonElement>('[data-fc-cat="fixed"]')
          ?.getAttribute('aria-selected'),
      ).toBe('true');
      expect(
        numberPanel.querySelector<HTMLInputElement>('input[type="number"][min="0"][max="10"]')
          ?.dataset.fcMixed,
      ).toBe(undefined);

      handle.open('fill');
      const fillGallery = document.querySelector<HTMLElement>('.fc-fmtdlg__fill-pattern-gallery');
      if (!fillGallery) throw new Error('fill pattern gallery missing');
      expect(
        [...fillGallery.querySelectorAll<HTMLButtonElement>('[data-fc-fill-pattern]')].every(
          (button) => button.getAttribute('aria-pressed') === 'false',
        ),
      ).toBe(true);
      const selectedPattern = fillGallery.querySelector<HTMLButtonElement>(
        '[data-fc-fill-pattern="darkGrid"]',
      );
      selectedPattern?.click();
      expect(selectedPattern?.getAttribute('aria-pressed')).toBe('true');

      handle.open('more');
      const morePanel = document.querySelector<HTMLElement>(
        '.fc-fmtdlg__panel-tab[data-fc-tab="more"]',
      );
      if (!morePanel) throw new Error('more panel missing');
      const validationSelects = morePanel.querySelectorAll<HTMLSelectElement>('select');
      expect(validationSelects[0]?.value).toBe('');
      expect(validationSelects[1]?.value).toBe('');
      expect(
        [...morePanel.querySelectorAll<HTMLInputElement>('input[type="radio"]')].every(
          (input) => !input.checked,
        ),
      ).toBe(true);
      expect(
        [...morePanel.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')].every(
          (input) => input.indeterminate,
        ),
      ).toBe(true);
      expect(
        [...morePanel.querySelectorAll<HTMLInputElement>('input[type="number"]')].every(
          (input) => input.value === '',
        ),
      ).toBe(true);
      const validationArea = [...morePanel.querySelectorAll<HTMLTextAreaElement>('textarea')].at(
        -1,
      );
      expect(validationArea?.value).toBe('');

      const validationKind = validationSelects[0];
      if (!validationKind) throw new Error('validation kind control missing');
      validationKind.value = 'list';
      validationKind.dispatchEvent(new Event('change', { bubbles: true }));
      expect(validationKind.value).toBe('list');
      expect(validationKind.dataset.fcMixed).toBe(undefined);
      expect(
        [...morePanel.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')].every(
          (input) => !input.indeterminate,
        ),
      ).toBe(true);
    } finally {
      handle.detach();
    }
  });
});
