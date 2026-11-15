import '../../../api/setup-dom.js';
import { afterEach, describe, expect, it } from 'vitest';
import { mount } from '../../../../src/index.js';
import { ribbon } from '../../../../src/toolbar/index.js';

afterEach(() => {
  document.body.replaceChildren();
});

// Home split buttons whose menu half the desktop opens with Alt+Down.
const HOME_SPLITS = ['paste', 'copy', 'underline', 'merge', 'currency', 'autosum'] as const;

const altDown = (): KeyboardEvent =>
  new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true, cancelable: true });

describe('ribbon split buttons', () => {
  it.each(['default', 'mac'] as const)(
    'open their menu with Alt+Down on the %s profile',
    async (platform) => {
      const host = document.createElement('div');
      document.body.appendChild(host);
      const instance = await mount(host, { locale: 'en', ui: { platform }, toolbar: ribbon() });
      try {
        host.querySelector<HTMLButtonElement>('[data-ribbon-tab="home"]')?.click();
        const closed: string[] = [];
        for (const id of HOME_SPLITS) {
          const button = host.querySelector<HTMLButtonElement>(
            `button[data-ribbon-command="${id}"][data-ribbon-menu-id]`,
          );
          if (!button) {
            closed.push(`${id}: not rendered as a split button`);
            continue;
          }
          expect(button.getAttribute('aria-haspopup')).toBe('menu');
          button.focus();
          const event = altDown();
          button.dispatchEvent(event);
          if (button.getAttribute('aria-expanded') !== 'true' || !event.defaultPrevented) {
            closed.push(id);
          }
          document.activeElement?.dispatchEvent(
            new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
          );
        }
        expect(closed).toEqual([]);
      } finally {
        instance.dispose();
      }
    },
  );
});
