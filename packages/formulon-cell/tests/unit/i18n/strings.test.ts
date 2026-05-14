import { describe, expect, it } from 'vitest';
import { defaultStrings, dictionaries, en, ja, mergeStrings } from '../../../src/i18n/strings.js';

type JsonRecord = Record<string, unknown>;

const isRecord = (value: unknown): value is JsonRecord =>
  value !== null && typeof value === 'object' && !Array.isArray(value);

function collectLeafPaths(value: unknown, prefix = ''): string[] {
  if (!isRecord(value)) return [prefix];
  return Object.keys(value)
    .sort()
    .flatMap((key) => collectLeafPaths(value[key], prefix ? `${prefix}.${key}` : key));
}

function collectStringLeaves(value: unknown, prefix = ''): [path: string, value: unknown][] {
  if (!isRecord(value)) return [[prefix, value]];
  return Object.keys(value)
    .sort()
    .flatMap((key) => collectStringLeaves(value[key], prefix ? `${prefix}.${key}` : key));
}

describe('dictionaries', () => {
  it('exposes ja and en under the dictionaries record', () => {
    expect(dictionaries.ja).toBe(ja);
    expect(dictionaries.en).toBe(en);
  });

  it('defaults to Japanese', () => {
    expect(defaultStrings).toBe(ja);
  });

  it('keeps the same shape for ja and en', () => {
    const jaSections = Object.keys(ja).sort();
    const enSections = Object.keys(en).sort();
    expect(jaSections).toEqual(enSections);
    for (const section of jaSections) {
      const jaKeys = Object.keys(
        (ja as unknown as Record<string, Record<string, string>>)[section] ?? {},
      ).sort();
      const enKeys = Object.keys(
        (en as unknown as Record<string, Record<string, string>>)[section] ?? {},
      ).sort();
      expect(jaKeys).toEqual(enKeys);
    }
  });

  it('keeps the same nested leaf keys for ja and en', () => {
    expect(collectLeafPaths(ja)).toEqual(collectLeafPaths(en));
  });

  it('does not ship missing or empty dictionary leaves', () => {
    for (const [locale, dict] of Object.entries(dictionaries)) {
      const badLeaves = collectStringLeaves(dict)
        .filter(([, value]) => typeof value !== 'string' || value.trim() === '')
        .map(([path, value]) => `${locale}.${path}=${JSON.stringify(value)}`);

      expect(badLeaves).toEqual([]);
    }
  });
});

describe('mergeStrings', () => {
  it('returns the base unchanged when overlay is missing', () => {
    expect(mergeStrings(en)).toBe(en);
  });

  it('overlays a single key without disturbing the rest of the section', () => {
    const out = mergeStrings(en, { contextMenu: { copy: 'Duplicate' } });
    expect(out.contextMenu.copy).toBe('Duplicate');
    expect(out.contextMenu.cut).toBe(en.contextMenu.cut);
  });

  it('does not mutate the base dictionary', () => {
    const before = en.contextMenu.copy;
    mergeStrings(en, { contextMenu: { copy: 'Changed' } });
    expect(en.contextMenu.copy).toBe(before);
  });

  it('skips sections whose overlay value is falsy', () => {
    const out = mergeStrings(en, {
      contextMenu: undefined,
      formatDialog: { ok: 'Apply' },
    });
    expect(out.contextMenu).toEqual(en.contextMenu);
    expect(out.formatDialog.ok).toBe('Apply');
  });
});
