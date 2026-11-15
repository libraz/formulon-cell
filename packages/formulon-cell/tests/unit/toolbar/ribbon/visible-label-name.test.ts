import '../../../api/setup-dom.js';
import { afterEach, describe, expect, it } from 'vitest';
import { mount } from '../../../../src/index.js';
import { ribbon } from '../../../../src/toolbar/index.js';

afterEach(() => {
  document.body.replaceChildren();
});

const fold = (text: string): string =>
  text
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .trim();

describe('ribbon command buttons with a visible label', () => {
  it.each(['en', 'ja'] as const)(
    'carry that label in their accessible name (%s)',
    async (locale) => {
      const host = document.createElement('div');
      document.body.appendChild(host);
      const instance = await mount(host, { locale, toolbar: ribbon() });
      try {
        const mismatches: string[] = [];
        const tabs = [...host.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')];
        expect(tabs.length).toBeGreaterThan(5);
        for (const tab of tabs) {
          tab.click();
          for (const button of host.querySelectorAll<HTMLButtonElement>(
            'button[data-ribbon-command]',
          )) {
            const visible = [...button.children]
              .filter((child) => child.tagName === 'SPAN')
              .map((child) => child.textContent ?? '')
              .join(' ');
            const name = button.getAttribute('aria-label');
            // A single non-digit character is not checked as visible text.
            if (!name || !visible.trim() || /^\D$/.test(visible.trim())) continue;
            if (!fold(name).includes(fold(visible))) {
              mismatches.push(`${button.dataset.ribbonCommand}: "${visible}" / "${name}"`);
            }
          }
        }
        expect([...new Set(mismatches)]).toEqual([]);
      } finally {
        instance.dispose();
      }
    },
  );
});
