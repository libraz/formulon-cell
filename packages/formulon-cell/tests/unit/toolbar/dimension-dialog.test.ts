import { afterEach, describe, expect, it } from 'vitest';

import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { showDimensionDialog } from '../../../src/toolbar/dialogs/dimension.js';

describe('showDimensionDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('returns a bounded numeric dimension value', async () => {
    const pending = showDimensionDialog({ axis: 'row', strings: en, initial: 24, max: 409 });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Row height:');
    const input = dialog?.querySelector<HTMLInputElement>('input[inputmode="decimal"]');
    expect(input?.min).toBe('0');
    expect(input?.max).toBe('409');
    expect(input?.step).toBe('1');
    if (!input) throw new Error('Expected dimension input.');
    input.value = '48';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(48);
  });

  it.each([
    ['row', '行の高さ', '行の高さ:'],
    ['col', '列の幅', '列の幅:'],
  ] as const)(
    'titles the %s dialog as Excel does, without the menu ellipsis',
    (axis, title, label) => {
      // Excel: window title '行の高さ' / '列の幅' over the label '行の高さ:' / '列の幅:'.
      void showDimensionDialog({ axis, strings: ja, initial: 14, max: 409 });
      const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
      expect(dialog?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(title);
      expect(dialog?.getAttribute('aria-label')).toBe(title);
      expect(dialog?.textContent).toContain(label);
      expect(dialog?.querySelector('.fc-fmtdlg__panel')?.classList).toContain(`fc-dimdlg--${axis}`);
      dialog
        ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)')
        ?.click();
    },
  );
});
