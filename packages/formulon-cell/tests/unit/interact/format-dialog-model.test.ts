import { describe, expect, it } from 'vitest';

import {
  COMMON_FONTS,
  CURRENCY_SYMBOLS,
  defaultCurrencySymbolFor,
  isHexColor,
  normalizeFormatLocale,
  patternPresetsFor,
  THEME_SWATCHES,
} from '../../../src/interact/format-dialog-model.js';

describe('interact/format-dialog-model', () => {
  describe('normalizeFormatLocale', () => {
    it('expands short tags to BCP-47', () => {
      expect(normalizeFormatLocale('ja')).toBe('ja-JP');
      expect(normalizeFormatLocale('en')).toBe('en-US');
    });

    it('passes through full BCP-47 tags', () => {
      expect(normalizeFormatLocale('fr-CA')).toBe('fr-CA');
      expect(normalizeFormatLocale('zh-Hant-TW')).toBe('zh-Hant-TW');
    });

    it('defaults to en-US for empty input', () => {
      expect(normalizeFormatLocale('')).toBe('en-US');
    });
  });

  describe('defaultCurrencySymbolFor', () => {
    it('uses ¥ for Japanese locales', () => {
      expect(defaultCurrencySymbolFor('ja')).toBe('¥');
      expect(defaultCurrencySymbolFor('ja-JP')).toBe('¥');
    });

    it('uses $ for everything else', () => {
      expect(defaultCurrencySymbolFor('en')).toBe('$');
      expect(defaultCurrencySymbolFor('fr-CA')).toBe('$');
      expect(defaultCurrencySymbolFor('')).toBe('$');
    });
  });

  describe('patternPresetsFor', () => {
    it('returns Japanese date/time presets for ja locales', () => {
      const presets = patternPresetsFor('ja');
      expect(presets.date[0]).toContain('"年"');
      expect(presets.custom).toContain('¥#,##0;[Red]-¥#,##0');
    });

    it('returns US-style presets for en locales', () => {
      const presets = patternPresetsFor('en');
      expect(presets.date).toContain('m/d/yyyy');
      expect(presets.custom).toContain('$#,##0;[Red]-$#,##0');
    });

    it('falls back to US presets for unknown locales', () => {
      const presets = patternPresetsFor('xx');
      expect(presets.date).toContain('m/d/yyyy');
    });

    it('exposes the number-pattern category keys for every locale', () => {
      for (const loc of ['en', 'ja', 'fr-CA', 'pt-BR']) {
        const p = patternPresetsFor(loc);
        expect(Object.keys(p).sort()).toEqual(['custom', 'date', 'datetime', 'special', 'time']);
        for (const arr of Object.values(p)) expect(arr.length).toBeGreaterThan(0);
      }
    });
  });

  describe('isHexColor', () => {
    it('accepts 6-digit #rrggbb in mixed case', () => {
      expect(isHexColor('#000000')).toBe(true);
      expect(isHexColor('#ffffff')).toBe(true);
      expect(isHexColor('#AaBbCc')).toBe(true);
    });

    it('rejects 3-digit shorthand, alpha, and bare hex', () => {
      expect(isHexColor('#fff')).toBe(false);
      expect(isHexColor('#ffffffff')).toBe(false);
      expect(isHexColor('ffffff')).toBe(false);
      expect(isHexColor('red')).toBe(false);
      expect(isHexColor('')).toBe(false);
    });
  });

  describe('canonical option lists', () => {
    it('exposes a non-empty COMMON_FONTS list', () => {
      expect(COMMON_FONTS).toContain('Helvetica');
      expect(COMMON_FONTS).toContain('monospace');
    });

    it('exposes the 4 baseline currency symbols', () => {
      expect(CURRENCY_SYMBOLS).toEqual(['$', '¥', '€', '£']);
    });

    it('exposes the 12 theme swatches as valid hex colors', () => {
      expect(THEME_SWATCHES.length).toBe(12);
      for (const s of THEME_SWATCHES) expect(isHexColor(s)).toBe(true);
    });
  });
});
