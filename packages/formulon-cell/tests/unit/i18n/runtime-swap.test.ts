import { describe, expect, it, vi } from 'vitest';

import { createI18nController } from '../../../src/i18n/controller.js';
import { dictionaries } from '../../../src/i18n/strings.js';

describe('i18n/controller — runtime swap', () => {
  it('starts at the requested locale and exposes its strings', () => {
    const en = createI18nController({ locale: 'en' });
    expect(en.locale).toBe('en');
    expect(en.strings).toBe(dictionaries.en);
  });

  it('resolves regional locale tags through their base language dictionary', () => {
    const enUS = createI18nController({ locale: 'en-US' });
    expect(enUS.locale).toBe('en-US');
    expect(enUS.strings).toBe(dictionaries.en);

    enUS.setLocale('ja-JP');
    expect(enUS.locale).toBe('ja-JP');
    expect(enUS.strings).toBe(dictionaries.ja);
  });

  it('defaults to "en" when no locale is provided', () => {
    const c = createI18nController();
    expect(c.locale).toBe('en');
    expect(c.strings).toBe(dictionaries.en);
  });

  it('on("change") reports locale swaps and overrides until its disposer runs', () => {
    const c = createI18nController();
    const fn = vi.fn();
    const off = c.on('change', fn);

    c.setLocale('ja');
    c.override('ja', { formatDialog: { title: 'Custom' } });
    expect(fn).toHaveBeenCalledTimes(2);
    expect(c.strings.formatDialog.title).toBe('Custom');

    off();
    off();
    c.setLocale('en');
    expect(fn).toHaveBeenCalledTimes(2);
  });

  it('setLocale flips the active dictionary and notifies subscribers', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    const unsub = c.on('change', fn);

    c.setLocale('ja');
    expect(c.locale).toBe('ja');
    expect(c.strings).toBe(dictionaries.ja);
    expect(fn).toHaveBeenCalledTimes(1);
    expect(fn).toHaveBeenCalledWith(dictionaries.ja);
    unsub();
  });

  it('setLocale to the current locale is a no-op (no notification)', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    c.on('change', fn);
    c.setLocale('en');
    expect(fn).not.toHaveBeenCalled();
  });

  it('override deep-merges an overlay onto the current locale and notifies', () => {
    const c = createI18nController({ locale: 'en' });
    const before = c.strings.formatDialog.title;
    const fn = vi.fn();
    c.on('change', fn);

    c.override('en', { formatDialog: { title: 'Custom Title' } });

    expect(c.strings.formatDialog.title).toBe('Custom Title');
    expect(c.strings.formatDialog.title).not.toBe(before);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('override on a non-current locale does not notify until setLocale activates it', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    c.on('change', fn);

    c.override('ja', { formatDialog: { title: 'カスタム' } });
    expect(fn).not.toHaveBeenCalled();

    c.setLocale('ja');
    // setLocale notifies once with the merged ja
    expect(fn).toHaveBeenCalledTimes(1);
    expect(c.strings.formatDialog.title).toBe('カスタム');
  });

  it('subsequent override calls compose on top of the prior overlay', () => {
    const c = createI18nController({ locale: 'en' });
    c.override('en', { formatDialog: { title: 'First' } });
    c.override('en', { formatDialog: { ok: 'Apply' } });
    expect(c.strings.formatDialog.title).toBe('First');
    expect(c.strings.formatDialog.ok).toBe('Apply');
  });

  it('on returns an unsubscribe that stops further notifications', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    const unsub = c.on('change', fn);
    unsub();
    c.setLocale('ja');
    expect(fn).not.toHaveBeenCalled();
  });

  it('dispose clears the listener registry — subsequent setLocale notifies nobody', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    c.on('change', fn);
    c.dispose();
    c.setLocale('ja');
    expect(fn).not.toHaveBeenCalled();
  });
});
