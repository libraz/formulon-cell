import { afterEach, describe, expect, it } from 'vitest';
import { en, ja } from '../../../../src/i18n/strings.js';
import { confirmMoveOverwrite } from '../../../../src/toolbar/dialogs/move-overwrite-confirm.js';

const workbook = (profile: string) => ({ spreadsheetProfileId: () => profile });
const message = () => document.body.querySelector('.fc-tb__dlg__message')?.textContent;
const click = (label: string): void => {
  [...document.body.querySelectorAll('button')]
    .find((b) => b.textContent === label)
    ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
};

afterEach(() => {
  document.body.innerHTML = '';
});

describe('confirmMoveOverwrite', () => {
  it('asks in the Mac wording, Cancel left of OK, and resolves the choice', async () => {
    // Excel (Mac, ja): 「既にデータがありますが、置き換えますか?」 キャンセル / OK (default).
    const pending = confirmMoveOverwrite(ja, workbook('mac-ja_JP'));
    await Promise.resolve();
    expect(message()).toBe('既にデータがありますが、置き換えますか?');
    const labels = [...document.body.querySelectorAll('button')]
      .map((b) => b.textContent)
      .filter((t) => t === 'キャンセル' || t === 'OK');
    expect(labels).toEqual(['キャンセル', 'OK']);
    click('キャンセル');
    expect(await pending).toBe(false);
  });

  it('uses the shared wording off the Mac profile', async () => {
    const pending = confirmMoveOverwrite(en, workbook('windows-en_US'));
    await Promise.resolve();
    expect(message()).toBe("There's already data here. Do you want to replace it?");
    click('OK');
    expect(await pending).toBe(true);
  });
});
