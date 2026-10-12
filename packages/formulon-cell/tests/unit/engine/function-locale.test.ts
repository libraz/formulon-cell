import { describe, expect, it } from 'vitest';
import {
  canonicalizeFormula,
  canonicalizeFormulaForEngine,
  localizeFormula,
  spreadsheetProfileForLocaleOrdinal,
} from '../../../src/engine/function-locale.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const wb = (
  alias: Record<string, string>,
  reverse: Record<string, string>,
  capability = true,
): WorkbookHandle =>
  ({
    capabilities: { functionLocale: capability },
    localizeFunctionName: (name: string) => alias[name] ?? name,
    canonicalizeFunctionName: (name: string) => reverse[name] ?? name,
  }) as unknown as WorkbookHandle;

describe('localizeFormula', () => {
  it('preserves the active host when adapting a legacy locale ordinal', () => {
    expect(spreadsheetProfileForLocaleOrdinal('windows-en_US', 2)).toBe('windows-de_DE');
    expect(spreadsheetProfileForLocaleOrdinal('mac-ja_JP', 3)).toBe('mac-fr_FR');
    expect(spreadsheetProfileForLocaleOrdinal(null, 1)).toBe('windows-ja_JP');
  });

  it('returns the formula unchanged for en-US (locale 0)', () => {
    const handle = wb({ SUM: '合計' }, {});
    expect(localizeFormula(handle, '=SUM(A1:A3)', 0)).toBe('=SUM(A1:A3)');
  });

  it('replaces canonical function names with their localized alias for non-en locales', () => {
    const handle = wb({ SUM: '合計', IF: 'もし' }, {});
    expect(localizeFormula(handle, '=SUM(A1)+IF(B1>0,1,0)', 1)).toBe('=合計(A1)+もし(B1>0,1,0)');
  });

  it('does not touch identifiers without a registered alias', () => {
    const handle = wb({}, {});
    expect(localizeFormula(handle, '=COUNT(A1:A3)', 1)).toBe('=COUNT(A1:A3)');
  });

  it('preserves contents of string literals verbatim', () => {
    const handle = wb({ SUM: '合計' }, {});
    expect(localizeFormula(handle, '=CONCAT("SUM(", A1, ")")', 1)).toBe('=CONCAT("SUM(", A1, ")")');
  });

  it('preserves quoted sheet names and structured-reference contents', () => {
    const handle = wb({ SUM: 'SUMME' }, {});
    expect(
      localizeFormula(
        handle,
        "='O''SUMME(A)'!B1+'SUMME(A)'!B1+SUM(Table1[SUMME(A)])+SUM(Table1[[#Headers],[SUMME(A)]])",
        2,
      ),
    ).toBe(
      "='O''SUMME(A)'!B1+'SUMME(A)'!B1+SUMME(Table1[SUMME(A)])+SUMME(Table1[[#Headers],[SUMME(A)]])",
    );
  });

  it('localizes function names after implicit-intersection markers', () => {
    const handle = wb({ SUM: 'SUMME' }, {});
    expect(localizeFormula(handle, '=@SUM(A1)', 2)).toBe('=@SUMME(A1)');
  });

  it('returns input unchanged when the engine lacks the capability', () => {
    const handle = wb({ SUM: '合計' }, {}, false);
    expect(localizeFormula(handle, '=SUM(A1)', 1)).toBe('=SUM(A1)');
  });

  it('uses German and French ordinals while canonical locales stay unchanged', () => {
    const calls: number[] = [];
    const handle = {
      capabilities: { functionLocale: true },
      localizeFunctionName: (name: string, locale: number) => {
        calls.push(locale);
        return locale === 2 ? 'SUMME' : locale === 3 ? 'SOMME' : name;
      },
    } as unknown as WorkbookHandle;
    expect(localizeFormula(handle, '=SUM(A1)', 2)).toBe('=SUMME(A1)');
    expect(localizeFormula(handle, '=SUM(A1)', 3)).toBe('=SOMME(A1)');
    expect(localizeFormula(handle, '=SUM(A1)', 0)).toBe('=SUM(A1)');
    expect(calls).toEqual([2, 3]);
  });
});

describe('canonicalizeFormula', () => {
  it('returns the formula unchanged for en-US (locale 0)', () => {
    const handle = wb({}, { 合計: 'SUM' });
    expect(canonicalizeFormula(handle, '=合計(A1:A3)', 0)).toBe('=合計(A1:A3)');
  });

  it('replaces localized names with their canonical form', () => {
    const handle = wb({}, { 合計: 'SUM', もし: 'IF' });
    expect(canonicalizeFormula(handle, '=合計(A1)+もし(B1>0,1,0)', 1)).toBe(
      '=SUM(A1)+IF(B1>0,1,0)',
    );
  });

  it('is idempotent on already-canonical formulas', () => {
    const handle = wb({}, { 合計: 'SUM' });
    expect(canonicalizeFormula(handle, '=SUM(A1)', 1)).toBe('=SUM(A1)');
  });

  it('preserves string literal contents', () => {
    const handle = wb({}, { 合計: 'SUM' });
    expect(canonicalizeFormula(handle, '=CONCAT("合計(", A1, ")")', 1)).toBe(
      '=CONCAT("合計(", A1, ")")',
    );
  });

  it('preserves quoted sheet names and structured-reference contents', () => {
    const handle = wb({}, { SUMME: 'SUM' });
    expect(
      canonicalizeFormula(
        handle,
        "='O''SUMME(A)'!B1+'SUMME(A)'!B1+SUMME(Table1[SUMME(A)])+SUMME(Table1[[#Headers],[SUMME(A)]])",
        2,
      ),
    ).toBe(
      "='O''SUMME(A)'!B1+'SUMME(A)'!B1+SUM(Table1[SUMME(A)])+SUM(Table1[[#Headers],[SUMME(A)]])",
    );
  });

  it('canonicalizes function names after implicit-intersection markers', () => {
    const handle = wb({}, { SUMME: 'SUM' });
    expect(canonicalizeFormula(handle, '=@SUMME(A1)', 2)).toBe('=@SUM(A1)');
  });

  it('preserves canonical names before resolving colliding locale aliases', () => {
    const handle = {
      capabilities: { functionLocale: true, functionMetadata: true },
      functionNames: () => ['DOLLAR', 'USDOLLAR'],
      canonicalizeFunctionName: (name: string) =>
        ({ DOLLAR: 'USDOLLAR', YEN: 'DOLLAR' })[name] ?? name,
    } as unknown as WorkbookHandle;

    expect(canonicalizeFormulaForEngine(handle, '=DOLLAR(1)+YEN(2)', 1)).toBe(
      '=DOLLAR(1)+DOLLAR(2)',
    );
    expect(canonicalizeFormula(handle, '=DOLLAR(1)', 1)).toBe('=USDOLLAR(1)');
  });
});
