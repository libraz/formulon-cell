import '../../../api/setup-dom.js';
import { afterEach, describe, expect, it } from 'vitest';
import { mount } from '../../../../src/index.js';
import { ribbon } from '../../../../src/toolbar/index.js';

afterEach(() => {
  document.body.replaceChildren();
});

const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 20));

// A dismissed ribbon dialog hands focus to the sheet, as the desktop app does.
const DIALOG_COMMANDS = [
  'fx',
  'namedRanges',
  'dataValidation',
  'protect',
  'textToColumns',
  'errorChecking',
] as const;

const openDialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('[role="dialog"]:not([hidden])');

const cancelButton = (dialog: HTMLElement): HTMLButtonElement | undefined =>
  [...dialog.querySelectorAll<HTMLButtonElement>('button')].find((button) =>
    /^(cancel|close)$/i.test(button.textContent?.trim() ?? ''),
  );

describe('ribbon dialogs', () => {
  it.each(['cancel', 'escape'] as const)(
    'return focus to the sheet when dismissed with %s',
    async (how) => {
      const host = document.createElement('div');
      document.body.appendChild(host);
      const instance = await mount(host, { locale: 'en', toolbar: ribbon() });
      try {
        const misplaced: string[] = [];
        const seen = new Set<string>();
        for (const tab of host.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')) {
          tab.click();
          for (const id of DIALOG_COMMANDS) {
            const button = host.querySelector<HTMLButtonElement>(
              `button[data-ribbon-command="${id}"]`,
            );
            if (!button || seen.has(id)) continue;
            seen.add(id);
            button.focus();
            button.click();
            await settle();
            const dialog = openDialog();
            if (!dialog) {
              misplaced.push(`${id}: no dialog`);
              continue;
            }
            if (how === 'cancel') cancelButton(dialog)?.click();
            else {
              (document.activeElement ?? dialog).dispatchEvent(
                new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
              );
            }
            await settle();
            const active = document.activeElement as HTMLElement | null;
            if (openDialog()) misplaced.push(`${id}: still open`);
            else if (active !== instance.host) {
              misplaced.push(`${id}: ${active?.tagName}.${active?.className}`);
            }
          }
        }
        expect(misplaced).toEqual([]);
        expect([...seen].sort()).toEqual([...DIALOG_COMMANDS].sort());
      } finally {
        instance.dispose();
      }
    },
  );
});
