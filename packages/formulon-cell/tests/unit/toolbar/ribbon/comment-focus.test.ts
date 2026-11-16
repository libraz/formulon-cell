import '../../../api/setup-dom.js';
import { afterEach, describe, expect, it } from 'vitest';
import { mount } from '../../../../src/index.js';
import { mountToolbar, ribbon } from '../../../../src/toolbar/index.js';

afterEach(() => {
  document.body.replaceChildren();
});

const frame = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 50));

const makeHost = (): HTMLElement => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  return host;
};

const placements = ['inside', 'outside'] as const;

describe('ribbon New Comment', () => {
  it.each(
    placements.flatMap((placement) =>
      (['en', 'ja'] as const).flatMap((locale) =>
        (['newCommentReview', 'commentInsert'] as const).map((id) => [placement, locale, id]),
      ),
    ),
  )(
    'focuses the editor and closes on Escape (%s ribbon, %s, %s)',
    async (placement, locale, id) => {
      const sheetHost = makeHost();
      const instance =
        placement === 'inside'
          ? await mount(sheetHost, { locale, toolbar: ribbon() })
          : await mount(sheetHost, { locale });
      const ribbonHost = placement === 'inside' ? sheetHost : makeHost();
      const toolbar = placement === 'outside' ? await mountToolbar(ribbonHost, instance) : null;
      try {
        ribbonHost.querySelector<HTMLButtonElement>('[data-ribbon-tab="review"]')?.click();
        const button = ribbonHost.querySelector<HTMLButtonElement>(
          `button[data-ribbon-command="${id}"]`,
        );
        expect(button, `button ${id}`).toBeTruthy();
        button?.click();
        await frame();
        const note = sheetHost.querySelector<HTMLElement>('.fc-cmtnote');
        expect(note?.hidden).toBe(false);
        expect(note?.contains(document.activeElement)).toBe(true);
        document.activeElement?.dispatchEvent(
          new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
        );
        await frame();
        expect(note?.hidden).toBe(true);
      } finally {
        toolbar?.dispose();
        instance.dispose();
      }
    },
  );
});
