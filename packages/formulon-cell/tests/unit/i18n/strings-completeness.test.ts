import { describe, expect, it } from 'vitest';

import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';

type AnyRec = Record<string, unknown>;

function flatKeys(obj: unknown, prefix = ''): string[] {
  if (obj === null || typeof obj !== 'object') return prefix ? [prefix] : [];
  const rec = obj as AnyRec;
  const out: string[] = [];
  for (const key of Object.keys(rec).sort()) {
    const path = prefix ? `${prefix}.${key}` : key;
    const value = rec[key];
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      out.push(...flatKeys(value, path));
    } else {
      out.push(path);
    }
  }
  return out;
}

describe('i18n/strings — locale parity', () => {
  const enKeys = flatKeys(en);
  const jaKeys = flatKeys(ja);

  it('en and ja have an identical key set', () => {
    const missingInJa = enKeys.filter((k) => !jaKeys.includes(k));
    const missingInEn = jaKeys.filter((k) => !enKeys.includes(k));
    expect(missingInJa, `keys present in en but not in ja: ${missingInJa.join(', ')}`).toEqual([]);
    expect(missingInEn, `keys present in ja but not in en: ${missingInEn.join(', ')}`).toEqual([]);
  });

  it('en and ja have non-empty string values for every key', () => {
    function flatEntries(obj: unknown, prefix = ''): [string, unknown][] {
      if (obj === null || typeof obj !== 'object') {
        return prefix ? [[prefix, obj]] : [];
      }
      const rec = obj as AnyRec;
      const out: [string, unknown][] = [];
      for (const key of Object.keys(rec)) {
        const path = prefix ? `${prefix}.${key}` : key;
        const value = rec[key];
        if (value && typeof value === 'object' && !Array.isArray(value)) {
          out.push(...flatEntries(value, path));
        } else {
          out.push([path, value]);
        }
      }
      return out;
    }

    const assertLeafValue = (locale: 'en' | 'ja', path: string, val: unknown): void => {
      if (Array.isArray(val)) {
        expect(val.length, `${locale}[${path}] should be a non-empty array`).toBeGreaterThan(0);
        for (const item of val) {
          expect(typeof item, `${locale}[${path}] items should be strings`).toBe('string');
          expect(
            (item as string).length,
            `${locale}[${path}] items should be non-empty`,
          ).toBeGreaterThan(0);
        }
        return;
      }
      expect(typeof val, `${locale}[${path}] should be string`).toBe('string');
      expect((val as string).length, `${locale}[${path}] should be non-empty`).toBeGreaterThan(0);
    };

    for (const [path, val] of flatEntries(en)) assertLeafValue('en', path, val);
    for (const [path, val] of flatEntries(ja)) assertLeafValue('ja', path, val);
  });
});
