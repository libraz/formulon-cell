import { describe, expect, it, vi } from 'vitest';
import {
  allFunctionNames,
  buildFunctionCatalog,
  FUNCTION_CATEGORY_NAMES,
  isFunctionUnavailableForInsertion,
  supportedFunctionNames,
} from '../../../src/commands/function-categories.js';
import { FUNCTION_SIGNATURES } from '../../../src/commands/refs.js';

describe('function category catalog', () => {
  it('keeps every function in exactly one family with stable counts', () => {
    const expectedCounts = {
      compatibility: 39,
      cube: 7,
      database: 12,
      datetime: 25,
      engineering: 54,
      financial: 55,
      information: 22,
      logical: 19,
      lookup: 40,
      math: 80,
      statistical: 111,
      text: 50,
      web: 3,
    } as const;
    const officialNames = Object.entries(expectedCounts).flatMap(([category]) => {
      const names = FUNCTION_CATEGORY_NAMES[category as keyof typeof FUNCTION_CATEGORY_NAMES];
      expect(names).toHaveLength(expectedCounts[category as keyof typeof expectedCounts]);
      expect(names).toEqual([...names].sort());
      expect(new Set(names).size).toBe(names.length);
      return names;
    });
    expect(officialNames).toHaveLength(517);
    expect(new Set(officialNames).size).toBe(517);
    expect(FUNCTION_CATEGORY_NAMES.math).toContain('ACOS');
    expect(FUNCTION_CATEGORY_NAMES.statistical).toEqual(
      expect.arrayContaining(['COUNT', 'COUNTIF', 'NORM.INV']),
    );
    expect(FUNCTION_CATEGORY_NAMES.compatibility).toContain('NORMINV');
    expect(FUNCTION_CATEGORY_NAMES.math).not.toContain('COUNT');
    expect(FUNCTION_CATEGORY_NAMES.math).not.toContain('COUNTIF');
    expect(FUNCTION_CATEGORY_NAMES.dynamicArray).toHaveLength(30);
    expect(FUNCTION_CATEGORY_NAMES.dynamicArray).toEqual(
      expect.arrayContaining(['FILTER', 'UNIQUE', 'SORT', 'LAMBDA', 'PERCENTOF']),
    );
    for (const unresolved of ['ANCHORARRAY', 'CEILING', 'COPILOT', 'PY', 'SINGLE', 'USDOLLAR']) {
      expect(officialNames).not.toContain(unresolved);
    }
  });

  it('only exposes names backed by a function signature', () => {
    for (const category of Object.keys(FUNCTION_CATEGORY_NAMES) as Array<
      keyof typeof FUNCTION_CATEGORY_NAMES
    >) {
      const names = supportedFunctionNames(category);
      expect(names).toEqual([...names].sort());
      for (const name of names) expect(name in FUNCTION_SIGNATURES).toBe(true);
    }
  });

  it('places well-known functions in their families', () => {
    expect(supportedFunctionNames('math')).toContain('SUM');
    expect(supportedFunctionNames('logical')).toContain('IF');
    expect(supportedFunctionNames('lookup')).toContain('VLOOKUP');
    expect(supportedFunctionNames('statistical')).toContain('COUNT');
    expect(supportedFunctionNames('datetime')).toContain('TODAY');
    expect(supportedFunctionNames('text')).toContain('CONCAT');
    expect(supportedFunctionNames('math')).not.toContain('COUNT');
    expect(supportedFunctionNames('logical')).not.toContain('SUM');
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
    expect(snapshot.entries.get('ACOS')?.argumentLabels).toEqual(
      FUNCTION_SIGNATURES.ACOS?.filter((label) => label !== '...') ?? [],
    );
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

  it('copies raw availability only for live catalogs and blocks exactly class 3', () => {
    const availability = new Map<string, number | undefined>([
      ['ACOS', 0],
      ['INFO', 1],
      ['CELL', 2],
      ['CUBEVALUE', 3],
      ['SUM', 99],
      ['COUNTIF', undefined],
    ]);
    const reader = {
      functionNames: () => [...availability.keys()],
      functionMetadata: (name: string) => ({
        name,
        minArity: 1,
        maxArity: 1,
        ...(availability.get(name) === undefined ? {} : { availability: availability.get(name) }),
      }),
    };
    const snapshot = buildFunctionCatalog(reader, 0);

    expect(snapshot.names).toEqual([...availability.keys()].sort());
    expect(snapshot.entries.get('ACOS')?.availability).toBe(0);
    expect(snapshot.entries.get('INFO')?.availability).toBe(1);
    expect(snapshot.entries.get('CELL')?.availability).toBe(2);
    expect(snapshot.entries.get('CUBEVALUE')?.availability).toBe(3);
    expect(snapshot.entries.get('SUM')?.availability).toBe(99);
    expect(snapshot.entries.get('COUNTIF')).not.toHaveProperty('availability');
    expect(isFunctionUnavailableForInsertion(undefined)).toBe(false);
    expect(isFunctionUnavailableForInsertion(0)).toBe(false);
    expect(isFunctionUnavailableForInsertion(1)).toBe(false);
    expect(isFunctionUnavailableForInsertion(2)).toBe(false);
    expect(isFunctionUnavailableForInsertion(99)).toBe(false);
    expect(isFunctionUnavailableForInsertion(3)).toBe(true);
  });

  it('does not infer availability from the static fallback catalog', () => {
    const functionMetadata = vi.fn(() => ({
      name: 'SUM',
      minArity: 1,
      maxArity: 1,
      availability: 3,
    }));
    const snapshot = buildFunctionCatalog({ functionNames: () => null, functionMetadata }, 0);

    expect(snapshot.source).toBe('fallback');
    expect(snapshot.entries.get('SUM')).not.toHaveProperty('availability');
  });

  it('intersects live names with explicit family membership only', () => {
    const liveNames = new Set(['SUM', 'ACOS', 'NOT_A_FAMILY_MEMBER']);
    expect(supportedFunctionNames('math', liveNames)).toEqual(['ACOS', 'SUM']);
    expect(supportedFunctionNames('logical', liveNames)).toEqual([]);
    expect(allFunctionNames(liveNames)).toEqual(['ACOS', 'NOT_A_FAMILY_MEMBER', 'SUM']);
  });

  it('keeps unresolved names available through All without assigning them to a family', () => {
    const liveNames = new Set(['ANCHORARRAY', 'CEILING', 'COPILOT', 'PY', 'SINGLE', 'USDOLLAR']);
    expect(allFunctionNames(liveNames)).toEqual([
      'ANCHORARRAY',
      'CEILING',
      'COPILOT',
      'PY',
      'SINGLE',
      'USDOLLAR',
    ]);
    for (const category of Object.keys(FUNCTION_CATEGORY_NAMES) as Array<
      keyof typeof FUNCTION_CATEGORY_NAMES
    >) {
      expect(supportedFunctionNames(category, liveNames)).toEqual([]);
    }
  });
});
