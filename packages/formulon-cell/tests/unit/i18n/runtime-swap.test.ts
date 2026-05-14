import { describe, expect, it, vi } from 'vitest';

import { createI18nController } from '../../../src/i18n/controller.js';
import { dictionaries } from '../../../src/i18n/strings.js';

describe('i18n/controller — runtime swap', () => {
  it('starts at the requested locale and exposes its strings', () => {
    const en = createI18nController({ locale: 'en' });
    expect(en.locale).toBe('en');
    expect(en.strings).toBe(dictionaries.en);
  });

  it('defaults to "ja" when no locale is provided', () => {
    const c = createI18nController();
    expect(c.locale).toBe('ja');
    expect(c.strings).toBe(dictionaries.ja);
  });

  it('setLocale flips the active dictionary and notifies subscribers', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    const unsub = c.subscribe(fn);

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
    c.subscribe(fn);
    c.setLocale('en');
    expect(fn).not.toHaveBeenCalled();
  });

  it('extend deep-merges an overlay onto the current locale and notifies', () => {
    const c = createI18nController({ locale: 'en' });
    const before = c.strings.formatDialog.title;
    const fn = vi.fn();
    c.subscribe(fn);

    c.extend('en', { formatDialog: { title: 'Custom Title' } });

    expect(c.strings.formatDialog.title).toBe('Custom Title');
    expect(c.strings.formatDialog.title).not.toBe(before);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('extend on a non-current locale does not notify until setLocale activates it', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    c.subscribe(fn);

    c.extend('ja', { formatDialog: { title: 'カスタム' } });
    expect(fn).not.toHaveBeenCalled();

    c.setLocale('ja');
    // setLocale notifies once with the merged ja
    expect(fn).toHaveBeenCalledTimes(1);
    expect(c.strings.formatDialog.title).toBe('カスタム');
  });

  it('subsequent extend calls compose on top of the prior overlay', () => {
    const c = createI18nController({ locale: 'en' });
    c.extend('en', { formatDialog: { title: 'First' } });
    c.extend('en', { formatDialog: { ok: 'Apply' } });
    expect(c.strings.formatDialog.title).toBe('First');
    expect(c.strings.formatDialog.ok).toBe('Apply');
  });

  it('register adds a brand-new locale and switches into it via setLocale', () => {
    const c = createI18nController({ locale: 'en' });
    const customStrings = {
      ...dictionaries.en,
      formatDialog: { ...dictionaries.en.formatDialog, title: 'Titre' },
    };
    c.register('fr', customStrings);
    c.setLocale('fr');
    expect(c.locale).toBe('fr');
    expect(c.strings.formatDialog.title).toBe('Titre');
  });

  it('register on the current locale recomputes and notifies', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    c.subscribe(fn);
    const replaced = {
      ...dictionaries.en,
      formatDialog: { ...dictionaries.en.formatDialog, title: 'replaced' },
    };
    c.register('en', replaced);
    expect(fn).toHaveBeenCalledTimes(1);
    expect(c.strings.formatDialog.title).toBe('replaced');
  });

  it('subscribe returns an unsubscribe that stops further notifications', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    const unsub = c.subscribe(fn);
    unsub();
    c.setLocale('ja');
    expect(fn).not.toHaveBeenCalled();
  });

  it('dispose clears the listener registry — subsequent setLocale notifies nobody', () => {
    const c = createI18nController({ locale: 'en' });
    const fn = vi.fn();
    c.subscribe(fn);
    c.dispose();
    c.setLocale('ja');
    expect(fn).not.toHaveBeenCalled();
  });
});
