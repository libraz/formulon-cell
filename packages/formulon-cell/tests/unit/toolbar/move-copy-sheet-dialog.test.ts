import { afterEach, describe, expect, it } from 'vitest';
import { showMoveCopySheetDialog } from '../../../src/toolbar/dialogs/move-copy-sheet.js';

// Excel for Mac (ja) 移動またはコピー: book popup, insert-before list, copy checkbox.
const strings = {
  title: '移動またはコピー',
  targetBook: '移動先ブック名:',
  before: '挿入先:',
  toEnd: '(末尾へ移動)',
  makeCopy: 'コピーを作成する',
  copyUnavailable: 'シートのコピーはまだ使用できません。',
  ok: 'OK',
  cancel: 'キャンセル',
};

const open = (current = 0) =>
  showMoveCopySheetDialog({
    strings,
    bookName: 'book.xlsx',
    sheets: ['Sheet1', 'Sheet2', 'Sheet3'],
    current,
  });
const q = <T extends HTMLElement>(sel: string): T => {
  const el = document.querySelector<T>(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el;
};

describe('Move or Copy sheet dialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('lists the sheets and an end entry, with the target book and copy checkbox', async () => {
    const done = open();
    expect(q('.fc-fmtdlg__header').textContent).toBe('移動またはコピー');
    expect(q('.fc-movecopy__book').textContent).toContain('book.xlsx');
    expect(
      Array.from(
        document.querySelectorAll('.fc-movecopy__list [role="option"]'),
        (o) => o.textContent,
      ),
    ).toEqual(['Sheet1', 'Sheet2', 'Sheet3', '(末尾へ移動)']);
    const copy = q<HTMLInputElement>('.fc-movecopy__copy input');
    expect(copy.checked).toBe(false);
    expect(copy.disabled).toBe(true);
    expect(q('.fc-movecopy__copy').textContent).toBe('コピーを作成する');
    expect(
      Array.from(document.querySelectorAll('.fc-fmtdlg__footer button'), (b) => b.textContent),
    ).toEqual(['キャンセル', 'OK']);
    q<HTMLButtonElement>('.fc-fmtdlg__footer button').click();
    expect(await done).toBeNull();
  });

  it('starts on the sheet after the current one, so OK alone moves nothing', async () => {
    const done = open(0);
    const selected = document.querySelector('.fc-movecopy__list [aria-selected="true"]');
    expect(selected?.textContent).toBe('Sheet2');
    q<HTMLButtonElement>('.fc-fmtdlg__footer .fc-fmtdlg__btn--primary').click();
    expect(await done).toEqual({ before: 1 });
  });

  it('returns the chosen insert position, the end entry meaning after the last sheet', async () => {
    const done = open(0);
    const options = document.querySelectorAll<HTMLElement>('.fc-movecopy__list [role="option"]');
    options[3]?.click();
    expect(options[3]?.getAttribute('aria-selected')).toBe('true');
    q<HTMLButtonElement>('.fc-fmtdlg__footer .fc-fmtdlg__btn--primary').click();
    expect(await done).toEqual({ before: 3 });
  });
});
