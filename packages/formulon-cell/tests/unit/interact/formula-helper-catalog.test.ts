import { describe, expect, it, vi } from 'vitest';
import { resolveFormulaHelperDetails } from '../../../src/interact/formula-helper-catalog.js';

describe('formula helper catalog', () => {
  it('queries only the active function using the live BCP-47 locale', () => {
    const getMetadata = vi.fn(() => null);
    const result = resolveFormulaHelperDetails('SUM', {
      localeTag: 'de-DE',
      listSeparator: ';',
      getMetadata,
      localizeName: (name: string) => name,
    });

    expect(getMetadata).toHaveBeenCalledOnce();
    expect(getMetadata).toHaveBeenCalledWith('SUM', 'de-DE');
    expect(result).toEqual({
      signature: 'SUM(number1; [number2]; ...)',
      description: 'Addiert die Argumente.',
    });
  });

  it('prefers localized engine metadata for signatures and descriptions', () => {
    const getMetadata = vi.fn((name: string, locale: string) => {
      expect(name).toBe('SUM');
      expect(locale).toBe('fr-FR');
      return {
        localizedName: 'SOMME',
        signatureTemplate: 'SOMME(nombre1, [nombre2])',
        description: 'Additionne les nombres.',
        minArity: 1,
        maxArity: 2,
      };
    });
    expect(
      resolveFormulaHelperDetails('SUM', {
        localeTag: 'fr-FR',
        listSeparator: ';',
        getMetadata,
        localizeName: () => 'SUM',
      }),
    ).toEqual({ signature: 'SOMME(nombre1; [nombre2])', description: 'Additionne les nombres.' });
  });

  it('uses registered custom arguments when the engine has no metadata', () => {
    expect(
      resolveFormulaHelperDetails('MYFUNC', {
        localeTag: 'en-US',
        listSeparator: ',',
        getMetadata: () => null,
        getCustomFunctionArgs: () => ['value', '[fallback]'],
        localizeName: (name: string) => name,
      }),
    ).toEqual({ signature: 'MYFUNC(value, [fallback])' });
  });

  it('preserves an unbounded maximum arity from engine metadata', () => {
    expect(
      resolveFormulaHelperDetails('MYFUNC', {
        localeTag: 'en-US',
        getMetadata: () => ({ minArity: 1, maxArity: null }),
        getCustomFunctionArgs: () => ['value'],
        localizeName: (name: string) => name,
      }),
    ).toEqual({ signature: 'MYFUNC(value, ...)' });
  });

  it('does not invent a custom-function label for a built-in without a static signature', () => {
    expect(
      resolveFormulaHelperDetails('BUILTINWITHOUTSTATICHELP', {
        localeTag: 'en-US',
        listSeparator: ',',
        getMetadata: () => null,
        localizeName: (name: string) => name,
      }),
    ).toBeUndefined();
  });
});
