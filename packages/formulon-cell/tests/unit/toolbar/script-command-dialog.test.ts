import { afterEach, describe, expect, it } from 'vitest';

import { showScriptCommandDialog } from '../../../src/toolbar/dialogs/script-command.js';

describe('showScriptCommandDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('renders a dedicated command picker and returns the selected script action', async () => {
    const pending = showScriptCommandDialog({
      title: 'Script',
      label: 'Command',
      options: [
        { value: 'uppercase', label: 'Uppercase' },
        { value: 'lowercase', label: 'Lowercase' },
        { value: 'trim', label: 'Trim whitespace' },
      ],
      initial: 'uppercase',
      okLabel: 'Run',
      cancelLabel: 'Cancel',
    });

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Script');
    expect(dialog?.textContent).toContain('Trim whitespace');
    const select = dialog?.querySelector<HTMLSelectElement>('[data-script-command-select]');
    expect(select?.value).toBe('uppercase');
    expect(select?.className).toBe('app__dlg__select');
    expect(
      Array.from(select?.options ?? []).map((option) => [option.value, option.textContent]),
    ).toEqual([
      ['uppercase', 'Uppercase'],
      ['lowercase', 'Lowercase'],
      ['trim', 'Trim whitespace'],
    ]);
    if (!select) throw new Error('Expected script command select.');
    select.value = 'trim';

    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toBe('trim');
  });
});
