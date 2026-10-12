import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import type { FormulonModule, Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const makeHandle = (supportsProfile: boolean): { wb: WorkbookHandle; setCalls: string[] } => {
  const setCalls: string[] = [];
  const raw = supportsProfile
    ? {
        excelProfileId: () => ({
          status: { ok: true, status: 0, message: '', context: '' },
          value: 'win-365-ja_JP',
        }),
        setExcelProfileId: (profileId: string) => {
          setCalls.push(profileId);
          return { ok: true, status: 0, message: '', context: '' };
        },
      }
    : {};
  const module = { versionString: () => 'test' } as unknown as FormulonModule;
  const Ctor = WorkbookHandle as unknown as new (
    module: FormulonModule,
    wb: Workbook,
  ) => WorkbookHandle;
  return { wb: new Ctor(module, raw as unknown as Workbook), setCalls };
};

describe('WorkbookHandle spreadsheet profile', () => {
  it('reads and writes the engine formula-behaviour profile when available', () => {
    const { wb, setCalls } = makeHandle(true);

    expect(wb.capabilities.spreadsheetProfile).toBe(true);
    expect(wb.spreadsheetProfileId()).toBe('windows-ja_JP');
    expect(wb.setSpreadsheetProfileId('mac-ja_JP')).toBe(true);
    expect(setCalls).toEqual(['mac-365-ja_JP']);
  });

  it('is a null/false no-op when the engine package is older', () => {
    const { wb, setCalls } = makeHandle(false);

    expect(wb.capabilities.spreadsheetProfile).toBe(false);
    expect(wb.spreadsheetProfileId()).toBeNull();
    expect(wb.setSpreadsheetProfileId('mac-ja_JP')).toBe(false);
    expect(setCalls).toEqual([]);
  });

  it('uses full formula locale APIs with public profile overrides', () => {
    const raw = {
      excelProfileId: () => ({
        status: { ok: true, status: 0, message: '', context: '' },
        value: 'win-365-de_DE',
      }),
      setExcelProfileId: () => ({ ok: true, status: 0, message: '', context: '' }),
      localizeFunctionName: (name: string, profile: string) => ({
        status: { ok: true },
        value: profile.endsWith('de_DE') && name === 'SUM' ? 'SUMME' : name,
      }),
      canonicalizeFunctionName: (name: string) => ({ status: { ok: true }, value: name }),
      localizeFormula: (formula: string, profile: string) => ({
        status: { ok: true },
        value: `${profile}:${formula.replace('SUM', 'SUMME').replace(',', ';')}`,
      }),
      canonicalizeFormula: (formula: string, profile: string) => ({
        status: { ok: true },
        value: `${profile}:${formula.replace('SUMME', 'SUM').replace(';', ',')}`,
      }),
      localeFacts: (profile: string) => ({
        status: { ok: profile.endsWith('de_DE') },
        facts: profile.endsWith('de_DE')
          ? ({ decimalSeparator: ',', groupSeparator: '.', listSeparator: ';' } as never)
          : null,
      }),
    };
    const module = { versionString: () => 'test' } as unknown as FormulonModule;
    const Ctor = WorkbookHandle as unknown as new (
      module: FormulonModule,
      wb: Workbook,
    ) => WorkbookHandle;
    const wb = new Ctor(module, raw as unknown as Workbook);

    expect(wb.capabilities.formulaLocale).toBe(true);
    expect(wb.localizeFormulaText('=SUM(1,2)')).toBe('win-365-de_DE:=SUMME(1;2)');
    expect(wb.getFormulaForDisplay('=SUM(1,2)')).toBe('win-365-de_DE:=SUMME(1;2)');
    expect(wb.localizeFormulaText('=SUM(1,2)', 'mac-fr_FR')).toBe('mac-365-fr_FR:=SUMME(1;2)');
    expect(wb.canonicalizeLocalizedFormula('=SUMME(1;2)')).toBe('win-365-de_DE:=SUM(1,2)');
    expect(wb.localeFacts()?.listSeparator).toBe(';');
    expect(wb.localeFacts('windows-en_US')).toBeNull();
  });

  it('round-trips full formula localization and facts for all fourteen profiles', async () => {
    const profiles = [
      'windows-ja_JP',
      'windows-en_US',
      'windows-de_DE',
      'windows-fr_FR',
      'windows-zh_CN',
      'windows-ko_KR',
      'windows-th_TH',
      'mac-ja_JP',
      'mac-en_US',
      'mac-de_DE',
      'mac-fr_FR',
      'mac-zh_CN',
      'mac-ko_KR',
      'mac-th_TH',
    ] as const;
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.formulaLocale).toBe(true);
      expect(wb.capabilities.localeFacts).toBe(true);
      for (const profile of profiles) {
        expect(wb.setSpreadsheetProfileId(profile)).toBe(true);
        const canonical = '=SUM(1.5,TRUE)';
        const localized = wb.localizeFormulaText(canonical, profile);
        expect(wb.canonicalizeLocalizedFormula(localized, profile)).toBe(canonical);
        const facts = wb.localeFacts(profile);
        expect(facts).not.toBeNull();
        expect(facts?.decimalSeparator).toMatch(/[.,]/);
        expect(facts?.errorNames.length).toBeGreaterThan(0);
      }
    } finally {
      wb.dispose();
    }
  });

  it('falls back safely when a full formula locale call fails', () => {
    const raw = {
      excelProfileId: () => ({
        status: { ok: true },
        value: 'win-365-de_DE',
      }),
      setExcelProfileId: () => ({ ok: true }),
      localizeFunctionName: (name: string, profile: string) => {
        const aliases = profile.endsWith('fr_FR')
          ? { IF: 'SI', SUM: 'SOMME' }
          : { IF: 'WENN', SUM: 'SUMME' };
        return { status: { ok: true }, value: aliases[name as 'IF' | 'SUM'] ?? name };
      },
      canonicalizeFunctionName: (name: string) => {
        const canonical = { SI: 'IF', SOMME: 'SUM', WENN: 'IF', SUMME: 'SUM' };
        return { status: { ok: true }, value: canonical[name as keyof typeof canonical] ?? name };
      },
      localizeFormula: () => {
        throw new Error('unsupported');
      },
      canonicalizeFormula: () => ({ status: { ok: false }, value: '' }),
      localeFacts: () => ({ status: { ok: false }, facts: null }),
    };
    const module = { versionString: () => 'test' } as unknown as FormulonModule;
    const Ctor = WorkbookHandle as unknown as new (
      module: FormulonModule,
      wb: Workbook,
    ) => WorkbookHandle;
    const wb = new Ctor(module, raw as unknown as Workbook);

    expect(wb.localizeFormulaText('=SUM(1,2)')).toBe('=SUMME(1,2)');
    expect(wb.getFormulaForDisplay('=SUM(1,2)')).toBe('=SUMME(1,2)');
    expect(wb.localizeFormulaText('=IF(1,SUM(1,2),"SUM(")', 'windows-fr_FR')).toBe(
      '=SI(1,SOMME(1,2),"SUM(")',
    );
    expect(wb.canonicalizeLocalizedFormula('=SUMME(1;2)')).toBe('=SUM(1;2)');
    expect(
      wb.canonicalizeLocalizedFormula('=WENN(1;SUMME(1\\SUMME(2);"SUMME(";Table1[SUMME(A)]);0)'),
    ).toBe('=IF(1;SUM(1\\SUM(2);"SUMME(";Table1[SUMME(A)]);0)');
    expect(wb.localeFacts()).toBeNull();
  });

  it('canonicalizes German aliases at the formula write boundary', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
      wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=SUMME(1,2)');
      expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1,2)');
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 3 });
      const evaluated = wb.evaluateFormulaText({ sheet: 0, row: 0, col: 0 }, '=SUMME(2,3)');
      expect(evaluated.status.ok).toBe(true);
      expect(evaluated.value.number).toBe(5);
      const adHoc = wb.evalFormula('=SUMME(4,5)');
      expect(adHoc.status.ok).toBe(true);
      expect(adHoc.value.number).toBe(9);
      expect(wb.getFormulaForDisplay('=SUM(1,2)')).toBe('=SUMME(1;2)');
      expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1,2)');
      wb.applyCellPatchAtomic([
        {
          addr: { sheet: 0, row: 0, col: 0 },
          value: { kind: 'number', value: 6 },
          formula: '=SUMME(1,2,3)',
        },
      ]);
      expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1,2,3)');
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 6 });
    } finally {
      wb.dispose();
    }
  });

  it('preserves Japanese canonical currency names across write, eval, patch, and history', async () => {
    const wb = await WorkbookHandle.createDefault();
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    wb.attachHistory(history);
    try {
      expect(wb.setSpreadsheetProfileId('windows-ja_JP')).toBe(true);
      expect(wb.getFormulaForDisplay('=DOLLAR(1234.5)')).toBe('=YEN(1234.5)');

      wb.setFormula(addr, '=DOLLAR(1234.5)');
      expect(wb.cellFormula(addr)).toBe('=DOLLAR(1234.5)');
      expect(wb.getValue(addr)).toEqual({ kind: 'text', value: '¥1,235' });

      const evaluated = wb.evaluateFormulaText(addr, '=DOLLAR(1234.5)');
      expect(evaluated.status.ok).toBe(true);
      expect(evaluated.value.text).toBe('¥1,235');

      wb.setFormula(addr, '=1+1');
      wb.applyCellPatchAtomic([
        {
          addr,
          value: { kind: 'text', value: '¥1,235' },
          formula: '=DOLLAR(1234.5)',
        },
      ]);
      expect(wb.cellFormula(addr)).toBe('=DOLLAR(1234.5)');
      expect(history.undo()).toBe(true);
      expect(wb.cellFormula(addr)).toBe('=DOLLAR(1234.5)');
      expect(history.redo()).toBe(true);
      expect(wb.cellFormula(addr)).toBe('=1+1');
    } finally {
      wb.dispose();
    }
  });
});
