import { afterEach, describe, expect, it, vi } from 'vitest';
import { en } from '../../../src/i18n/strings.js';
import {
  showTableStyleDialog,
  type TableStyleDialogOptions,
} from '../../../src/toolbar/dialogs/table-style.js';

const dialog = (): HTMLElement | null => document.body.querySelector<HTMLElement>('.fc-tb__dlg');
const ok = (): void =>
  dialog()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

const open = (over: Partial<TableStyleDialogOptions> = {}) =>
  showTableStyleDialog({
    kind: 'table',
    mode: 'new',
    mac: false,
    strings: en,
    initial: { name: 'Style', elements: {}, stripeSizes: {} },
    editFormat: vi.fn(),
    describe: () => '',
    ...over,
  });

describe('table style dialog result', () => {
  afterEach(() => document.body.replaceChildren());

  it('stays open with the message when the name is taken', async () => {
    const pending = open({ validateName: (name) => (name === 'Style' ? 'taken' : null) });
    ok();
    expect(dialog()?.querySelector('.fc-tb__dlg__error')?.textContent).toBe('taken');
    const input = dialog()?.querySelector<HTMLInputElement>('input[data-dialog-field="name"]');
    if (input) input.value = 'Other';
    ok();
    await expect(pending).resolves.toMatchObject({ name: 'Other' });
  });

  it('drops the stripe size of an element the style no longer has', async () => {
    const pending = open({
      initial: {
        name: 'Style',
        elements: { firstRowStripe: { fill: '#336699' } },
        stripeSizes: { firstRowStripe: 2, firstColumnStripe: 3, secondRowStripe: 4 },
      },
    });
    ok();
    await expect(pending).resolves.toMatchObject({ stripeSizes: { firstRowStripe: 2 } });
    expect((await pending)?.stripeSizes).toEqual({ firstRowStripe: 2 });
  });

  it('keeps the default flag the dialog opens with', async () => {
    const pending = open({
      kind: 'pivot',
      mode: 'modify',
      mac: true,
      initial: { name: 'Style', elements: {}, stripeSizes: {}, setAsDefault: true },
    });
    ok();
    await expect(pending).resolves.toMatchObject({ setAsDefault: true });
  });
});
