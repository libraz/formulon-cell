import { afterEach, describe, expect, it } from 'vitest';
import { showUnhideSheetDialog } from '../../../src/toolbar/dialogs/unhide-sheet.js';

// Excel for Mac (ja) 再表示: one list of the hidden sheets with OK and キャンセル.
const strings = { title: '再表示', list: '再表示するシート:', ok: 'OK', cancel: 'キャンセル' };
const q = <T extends HTMLElement>(sel: string): T => {
  const el = document.querySelector<T>(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el;
};

describe('Unhide sheet dialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('lists the hidden sheets once, with the first one selected', async () => {
    const done = showUnhideSheetDialog({ strings, sheets: ['B', 'D'] });
    expect(q('.fc-fmtdlg__header').textContent).toBe('再表示');
    const options = document.querySelectorAll<HTMLElement>('.fc-movecopy__list [role="option"]');
    expect(Array.from(options, (o) => o.textContent)).toEqual(['B', 'D']);
    expect(options[0]?.getAttribute('aria-selected')).toBe('true');
    expect(
      Array.from(document.querySelectorAll('.fc-fmtdlg__footer button'), (b) => b.textContent),
    ).toEqual(['キャンセル', 'OK']);
    q<HTMLButtonElement>('.fc-fmtdlg__footer button').click();
    expect(await done).toBeNull();
  });

  it('resolves to the position of the chosen sheet', async () => {
    const done = showUnhideSheetDialog({ strings, sheets: ['B', 'D'] });
    document.querySelectorAll<HTMLElement>('.fc-movecopy__list [role="option"]')[1]?.click();
    q<HTMLButtonElement>('.fc-fmtdlg__footer .fc-fmtdlg__btn--primary').click();
    expect(await done).toBe(1);
  });
});
