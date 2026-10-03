import { describe, expect, it } from 'vitest';
import { toolbarLangForLocale } from '../../../src/toolbar/ribbon/mac/locale.js';

describe('toolbarLangForLocale', () => {
  it('treats ja* tags as Japanese and everything else as English', () => {
    expect(toolbarLangForLocale('ja')).toBe('ja');
    expect(toolbarLangForLocale('ja-JP')).toBe('ja');
    expect(toolbarLangForLocale('JA-jp')).toBe('ja');
    expect(toolbarLangForLocale('en')).toBe('en');
    expect(toolbarLangForLocale('en-US')).toBe('en');
    expect(toolbarLangForLocale('fr-FR')).toBe('en');
  });
});
