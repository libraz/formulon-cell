import { afterEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../../src/i18n/strings.js';
import { confirmSheetDelete } from '../../../../src/toolbar/dialogs/sheet-delete-confirm.js';

const workbook = (profile: string, cells: unknown[]): WorkbookHandle =>
  ({
    spreadsheetProfileId: () => profile,
    cells: function* () {
      yield* cells;
    },
  }) as unknown as WorkbookHandle;

const DATA = [{ addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 1 } }];

const buttons = () => [...document.body.querySelectorAll('button')].map((b) => b.textContent);
const click = (label: string): void => {
  [...document.body.querySelectorAll('button')]
    .find((b) => b.textContent === label)
    ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
};

afterEach(() => {
  document.body.innerHTML = '';
});

describe('confirmSheetDelete', () => {
  it('asks on the Mac profile before deleting a sheet with data, Cancel left of Delete', async () => {
    // Excel (Mac, ja): 「このシートは完全に削除されます。続けますか?」 キャンセル / 削除 (default).
    const pending = confirmSheetDelete(ja, workbook('mac-ja_JP', DATA), 0);
    await Promise.resolve();
    expect(document.body.querySelector('.fc-tb__dlg__message')?.textContent).toBe(
      'このシートは完全に削除されます。続けますか?',
    );
    expect(buttons().filter((t) => t === 'キャンセル' || t === '削除')).toEqual([
      'キャンセル',
      '削除',
    ]);
    click('キャンセル');
    expect(await pending).toBe(false);

    const accepted = confirmSheetDelete(ja, workbook('mac-ja_JP', DATA), 0);
    await Promise.resolve();
    click('削除');
    expect(await accepted).toBe(true);
  });

  it('deletes an empty sheet, or any sheet off the Mac profile, without asking', async () => {
    expect(await confirmSheetDelete(ja, workbook('mac-ja_JP', []), 0)).toBe(true);
    expect(await confirmSheetDelete(en, workbook('windows-en_US', DATA), 0)).toBe(true);
    expect(document.body.querySelector('button')).toBeNull();
  });
});
