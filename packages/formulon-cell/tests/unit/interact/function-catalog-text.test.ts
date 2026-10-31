import { describe, expect, it } from 'vitest';
import type { FunctionCatalogEntry } from '../../../src/commands/function-categories.js';
import {
  type FunctionNameReader,
  functionListSeparator,
  nativeFunctionName,
  resolveCatalogFunctionName,
} from '../../../src/interact/function-catalog-text.js';

const catalogOf = (names: readonly string[]) => ({
  knownNames: new Set(names),
  entries: new Map(
    names.map((name) => [name, { canonicalName: name } as FunctionCatalogEntry] as const),
  ),
});

const disposed = (): never => {
  throw new Error('WorkbookHandle is disposed');
};

describe('interact/function-catalog-text name resolution', () => {
  it('prefers a canonical catalog name over any alias', () => {
    const reader: FunctionNameReader = {
      localizeFunctionName: (name) => (name === 'USDOLLAR' ? 'DOLLAR' : name),
      canonicalizeFunctionName: (name) => (name === 'DOLLAR' ? 'USDOLLAR' : name),
    };
    expect(resolveCatalogFunctionName(catalogOf(['DOLLAR', 'USDOLLAR']), reader, 'dollar', 1)).toBe(
      'DOLLAR',
    );
  });

  it('takes the engine alias table before the native-name scan', () => {
    const reader: FunctionNameReader = {
      localizeFunctionName: (name) => (name === 'COUNT' ? 'ANZAHL' : name),
      canonicalizeFunctionName: (name) => (name.toUpperCase() === 'ANZAHL' ? 'SUM' : null),
    };
    expect(resolveCatalogFunctionName(catalogOf(['COUNT', 'SUM']), reader, 'Anzahl', 2)).toBe(
      'SUM',
    );
  });

  it('falls back to the native-name scan when canonicalization throws or is missing', () => {
    const localizeFunctionName = (name: string) => (name === 'SUM' ? 'SUMME' : name);
    const catalog = catalogOf(['SUM']);
    expect(
      resolveCatalogFunctionName(
        catalog,
        { localizeFunctionName, canonicalizeFunctionName: disposed },
        'SUMME',
        2,
      ),
    ).toBe('SUM');
    expect(resolveCatalogFunctionName(catalog, { localizeFunctionName }, 'summe', 2)).toBe('SUM');
  });

  it('returns null without throwing when every reader call throws', () => {
    const reader: FunctionNameReader = {
      localizeFunctionName: disposed,
      canonicalizeFunctionName: disposed,
    };
    expect(resolveCatalogFunctionName(catalogOf(['SUM']), reader, 'SUMME', 2)).toBeNull();
    expect(resolveCatalogFunctionName(catalogOf(['SUM']), null, 'SUMME', 2)).toBeNull();
  });

  it('keeps the canonical name when localization is unmapped, empty or throws', () => {
    expect(nativeFunctionName({ localizeFunctionName: () => '' }, 'SUM', 2)).toBe('SUM');
    expect(nativeFunctionName({ localizeFunctionName: disposed }, 'SUM', 2)).toBe('SUM');
    expect(nativeFunctionName(null, 'SUM', 2)).toBe('SUM');
    expect(nativeFunctionName({ localizeFunctionName: () => 'SUMME' }, 'SUM', 2)).toBe('SUMME');
  });

  it('reads the semicolon separator and defaults to a comma when the reader throws', () => {
    expect(functionListSeparator({ localeFacts: () => ({ listSeparator: ';' }) })).toBe(';');
    expect(functionListSeparator({ localeFacts: () => ({ listSeparator: ',' }) })).toBe(',');
    expect(functionListSeparator({ localeFacts: disposed })).toBe(',');
    expect(functionListSeparator(null)).toBe(',');
  });
});
