import { describe, expect, it } from 'vitest';
import { createI18nController } from '../../../src/i18n/controller.js';
import { en } from '../../../src/i18n/strings.js';

describe('i18n controller extend', () => {
  it('merges a second partial overlay naming a section the first lacked', () => {
    const c = createI18nController({ locale: 'en' });
    const a = Object.keys(en)[0] as keyof typeof en;
    const b = Object.keys(en)[1] as keyof typeof en;
    const keyA = Object.keys(en[a] as object)[0] as string;
    const keyB = Object.keys(en[b] as object)[0] as string;
    c.extend('en', { [a]: { [keyA]: 'A!' } } as never);
    expect(() => c.extend('en', { [b]: { [keyB]: 'B!' } } as never)).not.toThrow();
    const s = c.strings as unknown as Record<string, Record<string, string>>;
    expect(s[a]?.[keyA]).toBe('A!');
    expect(s[b]?.[keyB]).toBe('B!');
  });
});
