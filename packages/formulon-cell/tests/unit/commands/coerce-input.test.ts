import { describe, expect, it, vi } from 'vitest';
import {
  canonicalizeLocalizedFormulaForInput,
  coerceInput,
  coerceInputForCell,
  inputLocaleForWorkbook,
  parseLocalizedNumericInput,
  writeCoerced,
  writeInput,
  writeInputValidated,
} from '../../../src/commands/coerce-input.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('coerceInput', () => {
  it('returns blank for empty / whitespace strings', () => {
    expect(coerceInput('')).toEqual({ kind: 'blank' });
    expect(coerceInput('   ')).toEqual({ kind: 'blank' });
    expect(coerceInput('\t\n')).toEqual({ kind: 'blank' });
  });

  it('returns formula when leading "=" is present', () => {
    expect(coerceInput('=A1+1')).toEqual({ kind: 'formula', text: '=A1+1' });
  });

  it('trims surrounding whitespace before classifying as formula', () => {
    expect(coerceInput('  =SUM(A1:A3)  ')).toEqual({ kind: 'formula', text: '=SUM(A1:A3)' });
  });

  it('returns booleans case-insensitively', () => {
    expect(coerceInput('TRUE')).toEqual({ kind: 'bool', value: true });
    expect(coerceInput('FALSE')).toEqual({ kind: 'bool', value: false });
    expect(coerceInput('true')).toEqual({ kind: 'bool', value: true });
  });

  it('accepts upstream localized boolean text only for its matching locale', () => {
    expect(coerceInput('WAHR', { locale: 'de-DE' })).toEqual({ kind: 'bool', value: true });
    expect(coerceInput('wahr', { locale: 'de-DE' })).toEqual({ kind: 'bool', value: true });
    expect(coerceInput('FALSCH', { locale: 'de-DE' })).toEqual({ kind: 'bool', value: false });
    expect(coerceInput('VRAI', { locale: 'fr-FR' })).toEqual({ kind: 'bool', value: true });
    expect(coerceInput('faux', { locale: 'fr-FR' })).toEqual({ kind: 'bool', value: false });

    for (const locale of ['en-US', 'ja-JP', 'zh-CN', 'ko-KR', 'th-TH']) {
      expect(coerceInput('WAHR', { locale })).toEqual({ kind: 'text', value: 'WAHR' });
      expect(coerceInput('FALSCH', { locale })).toEqual({ kind: 'text', value: 'FALSCH' });
      expect(coerceInput('VRAI', { locale })).toEqual({ kind: 'text', value: 'VRAI' });
      expect(coerceInput('FAUX', { locale })).toEqual({ kind: 'text', value: 'FAUX' });
    }
    expect(coerceInput('VRAI', { locale: 'de-DE' })).toEqual({ kind: 'text', value: 'VRAI' });
    expect(coerceInput('WAHR', { locale: 'fr-FR' })).toEqual({ kind: 'text', value: 'WAHR' });
  });

  it('exposes strict numeric parsing for dialog criteria without date, boolean, or formula coercion', () => {
    expect(parseLocalizedNumericInput('1,5', 'de-DE')).toBe(1.5);
    expect(parseLocalizedNumericInput('1,5', 'fr-FR')).toBe(1.5);
    for (const locale of ['en-US', 'ja-JP', 'zh-CN', 'ko-KR', 'th-TH']) {
      expect(parseLocalizedNumericInput('1.5', locale)).toBe(1.5);
    }
    expect(parseLocalizedNumericInput('1.23,5', 'de-DE')).toBeNull();
    expect(parseLocalizedNumericInput('', 'de-DE')).toBeNull();
    expect(parseLocalizedNumericInput('TRUE', 'en-US')).toBeNull();
    expect(parseLocalizedNumericInput('25.12.2024', 'de-DE')).toBeNull();
    expect(parseLocalizedNumericInput('=SUMME(1,5;2)', 'de-DE')).toBeNull();
    expect(parseLocalizedNumericInput('1e999', 'en-US')).toBeNull();
  });

  it('maps public workbook profiles to input locales and falls back for older engines', () => {
    const withProfile = (profile: string | null): WorkbookHandle =>
      ({ spreadsheetProfileId: () => profile }) as unknown as WorkbookHandle;
    expect(inputLocaleForWorkbook(withProfile('windows-de_DE'))).toBe('de-DE');
    expect(inputLocaleForWorkbook(withProfile('mac-fr_FR'))).toBe('fr-FR');
    expect(inputLocaleForWorkbook(withProfile('windows-th_TH'))).toBe('th-TH');
    expect(inputLocaleForWorkbook(withProfile('unknown'))).toBeUndefined();
    expect(inputLocaleForWorkbook(stubHandle())).toBeUndefined();
  });

  it.each([
    ['windows-de_DE', 'de-DE', '=SUMME(1,5;2)'],
    ['windows-fr_FR', 'fr-FR', '=SOMME(1,5;2)'],
  ] as const)(
    'canonicalizes a localized formula at the raw input boundary (%s)',
    async (profile, locale, raw) => {
      const wb = await WorkbookHandle.createDefault();
      const store = createSpreadsheetStore();
      try {
        expect(wb.isStub).toBe(false);
        expect(wb.capabilities.formulaLocale).toBe(true);
        expect(wb.capabilities.localeFacts).toBe(true);
        expect(wb.setSpreadsheetProfileId(profile)).toBe(true);

        const input = coerceInputForCell(store.getState(), addr, raw, locale, wb);
        expect(input).toEqual({ kind: 'formula', text: '=SUM(1.5,2)' });
        expect(canonicalizeLocalizedFormulaForInput(wb, raw, locale)).toBe('=SUM(1.5,2)');

        writeInput(wb, addr, raw, undefined, locale);
        expect(wb.cellFormula(addr)).toBe('=SUM(1.5,2)');
      } finally {
        wb.dispose();
      }
    },
  );

  it('does not apply the active profile to an explicitly custom locale', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
      const canonicalize = vi.spyOn(wb, 'canonicalizeLocalizedFormula');

      const input = coerceInputForCell(store.getState(), addr, '=SUMME(A1)', 'de-brand', wb);
      expect(input).toEqual({ kind: 'formula', text: '=SUMME(A1)' });
      expect(canonicalizeLocalizedFormulaForInput(wb, '=SUMME(A1)', 'de-brand')).toBe('=SUMME(A1)');
      expect(canonicalize).not.toHaveBeenCalled();
    } finally {
      wb.dispose();
    }
  });

  it('rejects impossible dates, leaving them as text', () => {
    expect(coerceInput('13/45/2024')).toEqual({ kind: 'text', value: '13/45/2024' });
    expect(coerceInput('2024-02-30')).toEqual({ kind: 'text', value: '2024-02-30' });
  });

  it('uses a leading apostrophe to force text input', () => {
    expect(coerceInput("'123")).toEqual({ kind: 'text', value: '123' });
    expect(coerceInput("'=A1")).toEqual({ kind: 'text', value: '=A1' });
  });

  it('can force nonblank input to text for preformatted Text cells', () => {
    expect(coerceInput('123', { forceText: true })).toEqual({ kind: 'text', value: '123' });
    expect(coerceInput('=A1', { forceText: true })).toEqual({ kind: 'text', value: '=A1' });
    expect(coerceInput('', { forceText: true })).toEqual({ kind: 'blank' });
  });

  it('preserves the original (non-trimmed) string for text values', () => {
    // Trim is only applied for classification; the text value keeps its original spacing
    // so users can place leading-space text.
    expect(coerceInput('  hello  ')).toEqual({ kind: 'text', value: '  hello  ' });
  });

  it('does not parse partially-numeric strings as numbers', () => {
    expect(coerceInput('12abc')).toEqual({ kind: 'text', value: '12abc' });
    expect(coerceInput('1.2.3')).toEqual({ kind: 'text', value: '1.2.3' });
  });
});

const stubHandle = () => {
  return {
    setBlank: vi.fn(),
    setFormula: vi.fn(),
    setNumber: vi.fn(),
    setBool: vi.fn(),
    setError: vi.fn(),
    setText: vi.fn(),
  } as unknown as WorkbookHandle & {
    setBlank: ReturnType<typeof vi.fn>;
    setFormula: ReturnType<typeof vi.fn>;
    setNumber: ReturnType<typeof vi.fn>;
    setBool: ReturnType<typeof vi.fn>;
    setError: ReturnType<typeof vi.fn>;
    setText: ReturnType<typeof vi.fn>;
  };
};

const addr: Addr = { sheet: 0, row: 0, col: 0 };

describe('writeCoerced', () => {
  it('dispatches each kind to the matching setter', () => {
    const wb = stubHandle();
    writeCoerced(wb, addr, { kind: 'blank' });
    writeCoerced(wb, addr, { kind: 'formula', text: '=1+1' });
    writeCoerced(wb, addr, { kind: 'number', value: 7 });
    writeCoerced(wb, addr, { kind: 'bool', value: true });
    writeCoerced(wb, addr, { kind: 'text', value: 'hi' });
    writeCoerced(wb, addr, { kind: 'error', code: 6, text: '#N/A' });
    expect(wb.setBlank).toHaveBeenCalledWith(addr);
    expect(wb.setFormula).toHaveBeenCalledWith(addr, '=1+1');
    expect(wb.setNumber).toHaveBeenCalledWith(addr, 7);
    expect(wb.setBool).toHaveBeenCalledWith(addr, true);
    expect(wb.setText).toHaveBeenCalledWith(addr, 'hi');
    expect(wb.setError).toHaveBeenCalledWith(addr, 6);
  });
});

describe('writeInput', () => {
  it('coerces and writes through the workbook handle', () => {
    const wb = stubHandle();
    writeInput(wb, addr, '=A1');
    writeInput(wb, addr, '');
    writeInput(wb, addr, '3.14');
    expect(wb.setFormula).toHaveBeenCalledWith(addr, '=A1');
    expect(wb.setBlank).toHaveBeenCalledWith(addr);
    expect(wb.setNumber).toHaveBeenCalledWith(addr, 3.14);
  });

  it('normalizes R1C1 formula input to A1 before writing to the workbook', () => {
    const wb = stubHandle();

    writeInput(wb, { sheet: 0, row: 3, col: 3 }, '=SUM(R[-2]C[-2]:R[-1]C[-1])');

    expect(wb.setFormula).toHaveBeenCalledWith({ sheet: 0, row: 3, col: 3 }, '=SUM(B2:C3)');
  });

  it('normalizes R1C1 formula input before workbook calculation', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
    wb.setNumber({ sheet: 0, row: 1, col: 2 }, 3);
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 4);
    wb.setNumber({ sheet: 0, row: 2, col: 2 }, 5);

    writeInput(wb, { sheet: 0, row: 3, col: 3 }, '=SUM(R[-2]C[-2]:R[-1]C[-1])');
    wb.recalc();

    expect(wb.cellFormula({ sheet: 0, row: 3, col: 3 })).toBe('=SUM(B2:C3)');
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'number', value: 14 });
  });

  it('applies the implicit number format to the store on typed input (H-35)', () => {
    const wb = stubHandle();
    const store = createSpreadsheetStore();
    writeInput(wb, addr, '10%', store);
    expect(wb.setNumber).toHaveBeenCalledWith(addr, 0.1);
    expect(store.getState().format.formats.get('0:0:0')?.numFmt).toEqual({
      kind: 'percent',
      decimals: 0,
    });
  });

  it('does not override an explicit non-General cell format (H-35)', () => {
    const wb = stubHandle();
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, addr, { numFmt: { kind: 'fixed', decimals: 3 } });
    writeInput(wb, addr, '10%', store);
    expect(store.getState().format.formats.get('0:0:0')?.numFmt).toEqual({
      kind: 'fixed',
      decimals: 3,
    });
  });

  it('writes numeric and formula-looking input as text when the cell format is Text', () => {
    const wb = stubHandle();
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, addr, { numFmt: { kind: 'text' } });

    writeInput(wb, addr, '00123', store);
    writeInput(wb, addr, '=A1', store);

    expect(wb.setText).toHaveBeenCalledWith(addr, '00123');
    expect(wb.setText).toHaveBeenCalledWith(addr, '=A1');
    expect(wb.setNumber).not.toHaveBeenCalled();
    expect(wb.setFormula).not.toHaveBeenCalled();
  });
});

describe('writeInputValidated', () => {
  it('blocks invalid stop-style input when error alerts are enabled', () => {
    const wb = stubHandle();

    const outcome = writeInputValidated(wb, addr, 'Closed', {
      kind: 'list',
      source: ['Open'],
      errorStyle: 'stop',
    });

    expect(outcome.ok).toBe(false);
    expect(wb.setText).not.toHaveBeenCalled();
  });

  it('writes invalid input silently and reports success when the error alert is disabled (H-30)', () => {
    const wb = stubHandle();

    const outcome = writeInputValidated(wb, addr, 'Closed', {
      kind: 'list',
      source: ['Open'],
      errorStyle: 'stop',
      showErrorMessage: false,
    });

    // No blocking dialog: the caller must see success so the editor closes.
    expect(outcome.ok).toBe(true);
    expect(wb.setText).toHaveBeenCalledWith(addr, 'Closed');
  });
});

describe('typed dates in a 1904 workbook', () => {
  const store1904 = (): ReturnType<typeof createSpreadsheetStore> => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({ ...s, ui: { ...s.ui, date1904: true } }));
    return store;
  };
  const addr: Addr = { sheet: 0, row: 0, col: 0 };

  it('reads the date system from the store', () => {
    expect(coerceInputForCell(store1904().getState(), addr, '2024-01-01')).toMatchObject({
      kind: 'number',
      value: 43830,
    });
    expect(
      coerceInputForCell(createSpreadsheetStore().getState(), addr, '2024-01-01'),
    ).toMatchObject({ kind: 'number', value: 45292 });
  });

  it('keeps dates before 1904 as text', () => {
    expect(coerceInputForCell(store1904().getState(), addr, '1903-12-31')).toEqual({
      kind: 'text',
      value: '1903-12-31',
    });
    expect(coerceInput('1903-12-31', { date1904: true }).kind).toBe('text');
    expect(coerceInput('1903-12-31').kind).toBe('number');
  });

  it('applies to month-name and CJK date spellings', () => {
    expect(coerceInput('1-Jan-2024', { date1904: true })).toMatchObject({ value: 43830 });
    expect(coerceInput('2024年1月1日', { locale: 'ja-JP', date1904: true })).toMatchObject({
      value: 43830,
    });
  });
});

describe('coerceInputForCell under Set precision as displayed', () => {
  it('rounds a typed number to the digits its format shows', () => {
    const store = createSpreadsheetStore();
    const addr: Addr = { sheet: 0, row: 0, col: 0 };
    mutators.setCellFormat(store, addr, { numFmt: { kind: 'fixed', decimals: 2 } });
    expect(coerceInputForCell(store.getState(), addr, '1.23456')).toEqual({
      kind: 'number',
      value: 1.23456,
    });
    store.setState((s) => ({ ...s, ui: { ...s.ui, precisionAsDisplayed: true } }));
    expect(coerceInputForCell(store.getState(), addr, '1.23456')).toEqual({
      kind: 'number',
      value: 1.23,
    });
    // A General cell keeps full precision.
    expect(coerceInputForCell(store.getState(), { ...addr, col: 1 }, '1.23456')).toEqual({
      kind: 'number',
      value: 1.23456,
    });
  });
});
