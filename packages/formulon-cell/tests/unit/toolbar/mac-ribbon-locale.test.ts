import { describe, expect, it } from 'vitest';
import { toolbarLangForLocale } from '../../../src/toolbar/ribbon/mac/locale.js';

describe('toolbarLangForLocale', () => {
  it('resolves all seven languages from regional tags', () => {
    expect(toolbarLangForLocale('ja')).toBe('ja');
    expect(toolbarLangForLocale('ja-JP')).toBe('ja');
    expect(toolbarLangForLocale('JA-jp')).toBe('ja');
    expect(toolbarLangForLocale('en')).toBe('en');
    expect(toolbarLangForLocale('en-US')).toBe('en');
    for (const locale of ['de', 'fr', 'zh', 'ko', 'th']) {
      expect(toolbarLangForLocale(`${locale}-XX`)).toBe(locale);
    }
    expect(toolbarLangForLocale('unknown')).toBe('ja');
  });
});
