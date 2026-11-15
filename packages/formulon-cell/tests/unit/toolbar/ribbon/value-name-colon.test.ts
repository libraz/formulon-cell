import '../../../api/setup-dom.js';
import { afterEach, describe, expect, it } from 'vitest';
import { mount } from '../../../../src/index.js';
import { ribbon } from '../../../../src/toolbar/index.js';
import { projectValueName } from '../../../../src/toolbar/menu-a11y.js';

afterEach(() => {
  document.body.replaceChildren();
});

describe('projectValueName', () => {
  it.each([
    ['Font Size', '11', 'Font Size: 11'],
    ['Width:', 'Automatic', 'Width: Automatic'],
    ['Largeur :', 'Automatique', 'Largeur : Automatique'],
    ['幅：', '自動', '幅： 自動'],
  ])('joins %j and %j without doubling a colon', (name, value, expected) => {
    const el = document.createElement('button');
    expect(projectValueName(el, value, name)).toBe(expected);
    expect(el.getAttribute('aria-label')).toBe(expected);
  });
});

const locales = ['en', 'ja', 'de', 'fr', 'ko', 'th', 'zh'] as const;
const cases = (['default', 'mac'] as const).flatMap((platform) =>
  locales.map((locale) => [platform, locale] as const),
);

describe('ribbon dropdowns showing a value', () => {
  it.each(cases)(
    'never name themselves with a doubled colon (%s, %s)',
    async (platform, locale) => {
      const host = document.createElement('div');
      document.body.appendChild(host);
      const instance = await mount(host, { locale, ui: { platform }, toolbar: ribbon() });
      try {
        const doubled: string[] = [];
        for (const tab of host.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')) {
          tab.click();
          for (const button of host.querySelectorAll<HTMLElement>('.fc-tb__rb-dd__btn')) {
            const name = button.getAttribute('aria-label') ?? '';
            if (/[:：]\s*[:：]/.test(name)) doubled.push(name);
          }
        }
        expect([...new Set(doubled)]).toEqual([]);
      } finally {
        instance.dispose();
      }
    },
  );
});
