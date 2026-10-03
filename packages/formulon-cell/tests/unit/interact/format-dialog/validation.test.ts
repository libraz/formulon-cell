import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { coerceInput } from '../../../../src/commands/coerce-input.js';
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

  it('uses the shared range picker for validation list range sources', () => {
    setRange(store, 0, 0, 4, 2);
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more');

      const validationKind = document.querySelector<HTMLSelectElement>('select[aria-label="種類"]');
      const rangeRadio = document.querySelector<HTMLInputElement>(
        'input[name="fc-validation-list-source"][value="range"]',
      );
      const rangePicker = document.querySelector<HTMLButtonElement>(
        '[data-range-picker="format-validation-list-range"]',
      );
      const rangeInput = rangePicker
        ?.closest('.fc-range-picker')
        ?.querySelector<HTMLInputElement>('input');
      expect(rangeInput).toBeTruthy();
      expect(rangePicker?.getAttribute('aria-label')).toBe('範囲の選択');

      if (!validationKind || !rangeRadio || !rangeInput) {
        throw new Error('validation range controls missing');
      }
      validationKind.value = 'list';
      validationKind.dispatchEvent(new Event('change', { bubbles: true }));
      rangeRadio.checked = true;
      rangeRadio.dispatchEvent(new Event('change', { bubbles: true }));

      rangePicker?.click();
      expect(rangePicker?.dataset.rangePickerActive).toBe('true');
      expect(
        document.querySelector('.fc-fmtdlg')?.classList.contains('fc-fmtdlg--range-picking'),
      ).toBe(true);
      setRange(store, 2, 1, 6, 1);
      expect(rangeInput.value).toBe('B3:B7');

      handle.close();
      expect(rangePicker?.dataset.rangePickerActive).toBe('false');
      expect(
        document.querySelector('.fc-fmtdlg')?.classList.contains('fc-fmtdlg--range-picking'),
      ).toBe(false);
    } finally {
      handle.detach();
    }
  });

  const serialOf = (raw: string): number => {
    const coerced = coerceInput(raw);
    if (coerced.kind !== 'number') throw new Error(`expected serial for ${raw}`);
    return coerced.value;
  };

  // The Number tab also carries a `種類` select, so pick the validation-kind
  // select by an option only it owns.
  const findValidationKindSelect = (): HTMLSelectElement => {
    const select = Array.from(
      document.querySelectorAll<HTMLSelectElement>('select[aria-label="種類"]'),
    ).find((s) => Array.from(s.options).some((o) => o.value === 'textLength'));
    if (!select) throw new Error('validation kind select missing');
    return select;
  };

  const findValidationDropdownToggle = (): HTMLInputElement => {
    const label = Array.from(document.querySelectorAll<HTMLLabelElement>('label')).find((el) =>
      el.textContent?.includes('セル内ドロップダウンを表示する'),
    );
    const input = label?.querySelector<HTMLInputElement>('input[type="checkbox"]');
    if (!input) throw new Error('validation dropdown toggle missing');
    return input;
  };

  it('lets list validation hide the in-cell dropdown', () => {
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more', { mode: 'dataValidation', focus: 'validation' });

      const kind = findValidationKindSelect();
      kind.value = 'list';
      kind.dispatchEvent(new Event('change', { bubbles: true }));

      const dropdownToggle = findValidationDropdownToggle();
      expect(dropdownToggle.parentElement?.hidden).toBe(false);
      expect(dropdownToggle.checked).toBe(true);
      dropdownToggle.checked = false;
      dropdownToggle.dispatchEvent(new Event('change', { bubbles: true }));

      const listArea = document.querySelector<HTMLTextAreaElement>(
        'textarea[aria-label="値を直接入力"]',
      );
      if (!listArea) throw new Error('validation list textarea missing');
      listArea.value = 'A\nB';
      listArea.dispatchEvent(new Event('input', { bubbles: true }));

      const ok = Array.from(
        document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
      ).find((b) => b.textContent === 'OK');
      ok?.click();

      const validation = store
        .getState()
        .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.validation;
      expect(validation).toEqual({ kind: 'list', source: ['A', 'B'], showDropdown: false });
    } finally {
      handle.detach();
    }
  });

  it('hydrates hidden in-cell dropdown state into the list validation checkbox', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['A', 'B'], showDropdown: false } },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more', { mode: 'dataValidation', focus: 'validation' });
      const dropdownToggle = findValidationDropdownToggle();
      expect(dropdownToggle.parentElement?.hidden).toBe(false);
      expect(dropdownToggle.checked).toBe(false);
    } finally {
      handle.detach();
    }
  });

  it('date validation exposes date pickers for bounds and stores serials', () => {
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more', { mode: 'dataValidation', focus: 'validation' });

      const kind = findValidationKindSelect();
      kind.value = 'date';
      kind.dispatchEvent(new Event('change', { bubbles: true }));

      const aInput = document.querySelector<HTMLInputElement>('input[aria-label="値"]');
      const bInput = document.querySelector<HTMLInputElement>('input[aria-label="上限値"]');
      if (!aInput || !bInput) throw new Error('bound inputs missing');
      // A raw serial field would read `number`; the date picker reads `date`.
      expect(aInput.type).toBe('date');
      expect(bInput.type).toBe('date');

      aInput.value = '2026-01-01';
      aInput.dispatchEvent(new Event('input', { bubbles: true }));
      bInput.value = '2026-12-31';
      bInput.dispatchEvent(new Event('input', { bubbles: true }));

      const ok = Array.from(
        document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
      ).find((b) => b.textContent === 'OK');
      ok?.click();

      const validation = store
        .getState()
        .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.validation;
      expect(validation).toMatchObject({
        kind: 'date',
        op: 'between',
        a: serialOf('2026-01-01'),
        b: serialOf('2026-12-31'),
      });
    } finally {
      handle.detach();
    }
  });

  it('time validation uses a time picker for bounds', () => {
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more', { mode: 'dataValidation', focus: 'validation' });

      const kind = findValidationKindSelect();
      const op = document.querySelector<HTMLSelectElement>('select[aria-label="条件"]');
      if (!op) throw new Error('validation op select missing');
      kind.value = 'time';
      kind.dispatchEvent(new Event('change', { bubbles: true }));
      op.value = '>';
      op.dispatchEvent(new Event('change', { bubbles: true }));

      const aInput = document.querySelector<HTMLInputElement>('input[aria-label="値"]');
      if (!aInput) throw new Error('bound input missing');
      expect(aInput.type).toBe('time');

      aInput.value = '09:30';
      aInput.dispatchEvent(new Event('input', { bubbles: true }));

      const ok = Array.from(
        document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
      ).find((b) => b.textContent === 'OK');
      ok?.click();

      const validation = store
        .getState()
        .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.validation;
      expect(validation).toMatchObject({ kind: 'time', op: '>', a: serialOf('09:30') });
    } finally {
      handle.detach();
    }
  });

  it('hydrates date validation bounds back into the date picker', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        validation: {
          kind: 'date',
          op: 'between',
          a: serialOf('2026-03-10'),
          b: serialOf('2026-03-20'),
        },
      },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more', { mode: 'dataValidation', focus: 'validation' });
      const aInput = document.querySelector<HTMLInputElement>('input[aria-label="値"]');
      const bInput = document.querySelector<HTMLInputElement>('input[aria-label="上限値"]');
      expect(aInput?.type).toBe('date');
      expect(aInput?.value).toBe('2026-03-10');
      expect(bInput?.value).toBe('2026-03-20');
    } finally {
      handle.detach();
    }
  });
});
