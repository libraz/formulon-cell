import { describe, expect, it, vi } from 'vitest';
import {
  allFunctionNames,
  buildFunctionCatalog,
  FUNCTION_CATEGORY_NAMES,
  supportedFunctionNames,
} from '../../../src/commands/function-categories.js';
import { FUNCTION_SIGNATURES } from '../../../src/commands/refs.js';

describe('function category catalog', () => {
  it('only exposes names backed by a function signature', () => {
    for (const category of Object.keys(FUNCTION_CATEGORY_NAMES) as Array<
      keyof typeof FUNCTION_CATEGORY_NAMES
    >) {
      for (const name of supportedFunctionNames(category))
        expect(name in FUNCTION_SIGNATURES).toBe(true);
      expect(supportedFunctionNames(category)).toEqual(
        FUNCTION_CATEGORY_NAMES[category].filter((name) => name in FUNCTION_SIGNATURES).sort(),
      );
    }
  });

  it('keeps all engine functions available through the All Functions picker', () => {
    expect(allFunctionNames()).toEqual(Object.keys(FUNCTION_SIGNATURES).sort());
    expect(allFunctionNames().length).toBeGreaterThan(80);
  });

  it('builds an authoritative live catalog with metadata and localized display names', () => {
    const functionNames = vi.fn(() => ['sum', 'ACOS', 'acos', 'ACCRINT']);
    const functionMetadata = vi.fn((name: string, locale: number) => {
      if (name === 'ACOS') {
        return {
          name,
          minArity: 1,
          maxArity: 1,
          localizedName: locale === 1 ? 'ARCCOS' : 'ACOS',
          signatureTemplate: 'ACOS(number)',
          description: locale === 1 ? '余弦の逆関数' : 'Returns the arccosine.',
        };
      }
      if (name === 'ACCRINT') return { name, minArity: 0, maxArity: null };
      return { name, minArity: 1, maxArity: null };
    });

    const snapshot = buildFunctionCatalog({ functionNames, functionMetadata }, 1);

    expect(snapshot.source).toBe('engine');
    expect(snapshot.names).toEqual(['ACCRINT', 'ACOS', 'SUM']);
    expect(snapshot.knownNames).toEqual(new Set(['ACCRINT', 'ACOS', 'SUM']));
    expect(snapshot.entries.get('ACOS')).toMatchObject({
      canonicalName: 'ACOS',
      displayName: 'ARCCOS',
      minArity: 1,
      maxArity: 1,
      signatureTemplate: 'ACOS(number)',
      description: '余弦の逆関数',
    });
    expect(snapshot.entries.get('ACOS')?.argumentLabels).toEqual([]);
    expect(snapshot.entries.get('SUM')?.argumentLabels).toEqual(
      FUNCTION_SIGNATURES.SUM?.filter((label) => label !== '...'),
    );
    expect(snapshot.entries.get('ACCRINT')).toMatchObject({ minArity: 0, maxArity: null });
    expect(functionNames).toHaveBeenCalledTimes(1);
    expect(functionMetadata).toHaveBeenCalledTimes(3);
  });

  it('treats a valid empty engine catalog as authoritative and never guesses families', () => {
    const snapshot = buildFunctionCatalog(
      {
        functionNames: () => [],
        functionMetadata: vi.fn(),
      },
      0,
    );

    expect(snapshot).toMatchObject({ source: 'engine', names: [] });
    expect(snapshot.entries.size).toBe(0);
    expect(supportedFunctionNames('math', snapshot.knownNames)).toEqual([]);
    expect(allFunctionNames(snapshot.knownNames)).toEqual([]);
  });

  it('intersects live names with explicit family membership only', () => {
    const liveNames = new Set(['SUM', 'ACOS', 'NOT_A_FAMILY_MEMBER']);
    expect(supportedFunctionNames('math', liveNames)).toContain('SUM');
    expect(supportedFunctionNames('math', liveNames)).not.toContain('ACOS');
    expect(supportedFunctionNames('logical', liveNames)).toEqual([]);
    expect(allFunctionNames(liveNames)).toEqual(['ACOS', 'NOT_A_FAMILY_MEMBER', 'SUM']);
  });
});
