import { describe, expect, it } from 'vitest';
import type { FunctionCatalogEntry } from '../../../../src/commands/function-categories.js';
import {
  argumentFieldCount,
  insertableEntry,
  pickerSections,
  resolveArgumentHelp,
  sameAvailability,
} from '../../../../src/interact/mac-formula-palette-catalog.js';

const entry = (
  canonicalName: string,
  extra: Partial<FunctionCatalogEntry> = {},
): FunctionCatalogEntry => ({
  canonicalName,
  displayName: canonicalName,
  minArity: 1,
  maxArity: 3,
  argumentLabels: [],
  ...extra,
});

const entries = [
  entry('AND'),
  entry('IF', { argumentLabels: ['[test]', 'then'] }),
  entry('SUM'),
  entry('SUMIF', { availability: 3 }),
];

const catalog = {
  names: entries.map((e) => e.canonicalName),
  knownNames: new Set(entries.map((e) => e.canonicalName)),
  entries: new Map(entries.map((e) => [e.canonicalName, e])),
};

describe('mac-formula-palette-catalog insertableEntry', () => {
  it('returns the entry for a known, available function', () => {
    expect(insertableEntry(catalog, 'SUM')).toBe(entries[2]);
  });

  it('returns null for an unknown name', () => {
    expect(insertableEntry(catalog, 'NOPE')).toBeNull();
  });

  it('returns null for a function the engine marks unavailable', () => {
    expect(insertableEntry(catalog, 'SUMIF')).toBeNull();
  });

  it('returns null when the name is in entries but not in knownNames', () => {
    const narrowed = { ...catalog, knownNames: new Set(['AND']) };
    expect(insertableEntry(narrowed, 'SUM')).toBeNull();
  });
});

describe('mac-formula-palette-catalog pickerSections', () => {
  type Category = 'all' | 'recent' | 'logical';
  const recent = ['SUM', 'IF'];
  // Rows pair every category with every query and both recent states.
  const cases: {
    category: Category;
    query: string;
    withRecent: boolean;
    expected: [string, string[]][];
  }[] = [
    {
      category: 'all',
      query: '',
      withRecent: true,
      expected: [
        ['recent', ['SUM', 'IF']],
        ['all', ['AND', 'IF', 'SUM', 'SUMIF']],
      ],
    },
    {
      category: 'all',
      query: ' su ',
      withRecent: false,
      expected: [
        ['recent', []],
        ['all', ['SUM', 'SUMIF']],
      ],
    },
    {
      category: 'all',
      query: 'zz',
      withRecent: true,
      expected: [
        ['recent', []],
        ['all', []],
      ],
    },
    { category: 'recent', query: '', withRecent: false, expected: [['recent', []]] },
    { category: 'recent', query: 'su', withRecent: true, expected: [['recent', ['SUM']]] },
    { category: 'recent', query: 'zz', withRecent: false, expected: [['recent', []]] },
    { category: 'logical', query: '', withRecent: true, expected: [['logical', ['AND', 'IF']]] },
    { category: 'logical', query: 'a', withRecent: false, expected: [['logical', ['AND']]] },
    { category: 'logical', query: 'zz', withRecent: true, expected: [['logical', []]] },
  ];

  it.each(cases)('$category / "$query" / recent=$withRecent', (row) => {
    const sections = pickerSections(row.category, catalog, row.withRecent ? recent : [], row.query);
    expect(sections.map((s) => [s.key, s.names])).toEqual(row.expected);
  });
});

describe('mac-formula-palette-catalog resolveArgumentHelp', () => {
  const iff = entries[1] as FunctionCatalogEntry;

  it('strips optional brackets from the catalog label when no provider is given', () => {
    expect(resolveArgumentHelp(undefined, iff, 0, 'en-US', 'Argument')).toEqual({
      label: 'test',
      description: undefined,
      url: undefined,
    });
  });

  it('falls back to a numbered label past the catalog labels', () => {
    expect(resolveArgumentHelp(undefined, iff, 2, 'en-US', 'Argument').label).toBe('Argument 3');
  });

  it('lets provider fields override the catalog and passes name, index and locale', () => {
    const calls: unknown[][] = [];
    const help = resolveArgumentHelp(
      (...args) => {
        calls.push(args);
        return { label: 'Condition', description: 'Tested value', url: 'https://example.test' };
      },
      iff,
      1,
      'ja-JP',
      'Argument',
    );
    expect(calls).toEqual([['IF', 1, 'ja-JP']]);
    expect(help).toEqual({
      label: 'Condition',
      description: 'Tested value',
      url: 'https://example.test',
    });
  });

  it('keeps the catalog label when the provider omits it', () => {
    const help = resolveArgumentHelp(() => ({ description: 'only text' }), iff, 1, 'en-US', 'A');
    expect(help.label).toBe('then');
    expect(help.description).toBe('only text');
  });
});

describe('mac-formula-palette-catalog argument counts and availability', () => {
  it('shows the largest of min arity, current arguments and catalog labels', () => {
    const iff = entries[1] as FunctionCatalogEntry;
    expect(argumentFieldCount(iff, 0)).toBe(2);
    expect(argumentFieldCount(iff, 5)).toBe(5);
    expect(argumentFieldCount(entry('X', { minArity: 3 }), 1)).toBe(3);
  });

  it('compares availability with Object.is', () => {
    expect(sameAvailability(entry('A'), entry('B'))).toBe(true);
    expect(sameAvailability(entry('A', { availability: 3 }), entry('B'))).toBe(false);
  });
});
