import { describe, expect, it } from 'vitest';
import {
  displayLocaleTag,
  localeOrdinalForTag,
  localeTag,
  mergeFunctionMetadata,
} from '../../../src/engine/function-metadata.js';
import type { FunctionMetadataEntry, FunctionMetadataResult } from '../../../src/engine/types.js';

const base: FunctionMetadataResult = {
  ok: true,
  name: 'SUM',
  minArity: 1,
  maxArity: null,
  signatureTemplate: 'SUM(number1, [number2], ...)',
  description: 'Adds its arguments.',
};

describe('localeTag', () => {
  it('maps ordinals to BCP-47 tags and defaults to en-US', () => {
    expect(localeTag(0)).toBe('en-US');
    expect(localeTag(1)).toBe('ja-JP');
    expect(localeTag(2)).toBe('de-DE');
    expect(localeTag(3)).toBe('fr-FR');
    expect(localeTag(99)).toBe('en-US');
  });
});

describe('display locale tags', () => {
  it('normalizes every built-in UI locale without expanding native ordinals', () => {
    expect(['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'].map(displayLocaleTag)).toEqual([
      'en-US',
      'ja-JP',
      'de-DE',
      'fr-FR',
      'zh-CN',
      'ko-KR',
      'th-TH',
    ]);
    expect(localeOrdinalForTag('en-US')).toBe(0);
    expect(localeOrdinalForTag('ja-JP')).toBe(1);
    expect(localeOrdinalForTag('de-DE')).toBe(2);
    expect(localeOrdinalForTag('fr-FR')).toBe(3);
    expect(localeOrdinalForTag('zh-CN')).toBeNull();
    expect(localeOrdinalForTag('ko-KR')).toBeNull();
    expect(localeOrdinalForTag('th-TH')).toBeNull();
  });
});

describe('mergeFunctionMetadata', () => {
  it('returns the base verbatim when no provider entry is supplied', () => {
    expect(mergeFunctionMetadata(base, undefined, 'ja-JP')).toBe(base);
  });

  it('prefers per-locale overrides over entry defaults over engine values', () => {
    const entry: FunctionMetadataEntry = {
      signature: 'SUM(default)',
      description: 'default description',
      aliases: { 'ja-JP': '合計' },
      localized: {
        'ja-JP': { signature: 'SUM(数値1, ...)', description: '引数を合計します。' },
      },
    };
    expect(mergeFunctionMetadata(base, entry, 'ja-JP')).toEqual({
      ...base,
      signatureTemplate: 'SUM(数値1, ...)',
      description: '引数を合計します。',
      localizedName: '合計',
    });
  });

  it('falls back to entry defaults when the locale has no override', () => {
    const entry: FunctionMetadataEntry = {
      signature: 'SUM(default)',
      description: 'default description',
      localized: { 'ja-JP': { signature: 'SUM(数値1, ...)' } },
    };
    // en-US has no localized block, so entry-level defaults win over the base.
    expect(mergeFunctionMetadata(base, entry, 'en-US')).toEqual({
      ...base,
      signatureTemplate: 'SUM(default)',
      description: 'default description',
      localizedName: 'SUM',
    });
  });

  it('keeps engine values when the entry omits a field', () => {
    const entry: FunctionMetadataEntry = { aliases: { 'ja-JP': '合計' } };
    expect(mergeFunctionMetadata(base, entry, 'ja-JP')).toEqual({
      ...base,
      localizedName: '合計',
    });
  });

  it('uses the canonical name as localizedName when no alias matches the locale', () => {
    const entry: FunctionMetadataEntry = { aliases: { 'fr-FR': 'SOMME' } };
    expect(mergeFunctionMetadata(base, entry, 'ja-JP').localizedName).toBe('SUM');
  });
});
