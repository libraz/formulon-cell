import { describe, expect, it } from 'vitest';
import { canonicalizeFormula, localizeFormula } from '../../../src/engine/function-locale.js';
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

  it('returns input unchanged when the engine lacks the capability', () => {
    const handle = wb({ SUM: '合計' }, {}, false);
    expect(localizeFormula(handle, '=SUM(A1)', 1)).toBe('=SUM(A1)');
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
});
