import { afterEach, describe, expect, it } from 'vitest';
import { createDefaultRibbonHooks } from '../../../src/mount/toolbar-defaults.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { seedText } from '../mount/toolbar/fixtures.js';

const dialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-spelling:not([hidden])');
const buttons = (): HTMLButtonElement[] => [
  ...(dialog()?.querySelectorAll<HTMLButtonElement>('.fc-spelling__buttons button') ?? []),
];
const button = (label: string): HTMLButtonElement => {
  const found = buttons().find((candidate) => candidate.textContent === label);
  if (!found) throw new Error(`no ${label} button`);
  return found;
};
const disabled = (): string[] =>
  buttons()
    .filter((b) => b.disabled || b.getAttribute('aria-disabled') === 'true')
    .map((b) => b.textContent ?? '');
const checkbox = (label: string): HTMLInputElement => {
  const row = [...(dialog()?.querySelectorAll('.fc-spelling__check') ?? [])].find(
    (candidate) => candidate.textContent === label,
  );
  return row?.querySelector('input') as HTMLInputElement;
};
const labels = (): string[] =>
  [...(dialog()?.querySelectorAll<HTMLElement>('.fc-spelling__label') ?? [])]
    .filter((label) => !label.hidden)
    .map((label) => label.textContent ?? '');

// The Spelling dialog as the desktop app shows it on the Mac in Japanese.
describe('Spelling dialog controls on the Mac desktop', () => {
  let sheet: MountedStubSheet;
  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  const open = async (cells: readonly [number, number, string][]): Promise<void> => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    for (const [row, col, text] of cells) seedText(sheet, row, col, text);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    createDefaultRibbonHooks(sheet.instance).review?.spelling?.();
  };
  const text = (row: number, col: number): unknown =>
    sheet.workbook.getValue({ sheet: 0, row, col });

  it('checks English, labels the word inline and lists the dictionaries', async () => {
    await open([[0, 0, 'teh']]);
    expect(dialog()?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(
      'スペル チェック :English',
    );
    expect(labels()).toEqual(['辞書にない単語: teh', '修正候補:', '修正候補:', '追加先の辞書:']);
    expect(dialog()?.querySelector<HTMLSelectElement>('.fc-spelling__dictionary')).not.toBeNull();
  });

  it('starts with Always suggest on and Ignore UPPERCASE off, Suggest greyed', async () => {
    await open([[0, 0, 'teh']]);
    expect(checkbox('常に修正候補を表示する').checked).toBe(true);
    expect(checkbox('すべて大文字の単語は無視する').checked).toBe(false);
    expect(disabled()).toEqual(['候補表示', '直前の修正を元に戻す']);
  });

  it('makes Ignore Once the default button, and Return ignores the word once', async () => {
    await open([
      [0, 0, 'teh'],
      [1, 0, 'recieve'],
    ]);
    expect(button('無視').classList.contains('fc-fmtdlg__btn--primary')).toBe(true);
    expect(button('修正').classList.contains('fc-fmtdlg__btn--primary')).toBe(false);
    dialog()
      ?.querySelector('#fc-spelling-change-to')
      ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'teh' });
    expect(labels()[0]).toBe('辞書にない単語: recieve');
  });

  it('presses Ignore Once on Return when focus rests on the dialog itself', async () => {
    await open([
      [0, 0, 'teh'],
      [1, 0, 'recieve'],
    ]);
    const panel = dialog()?.querySelector<HTMLElement>('.fc-spelling__panel');
    panel?.focus();
    panel?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(labels()[0]).toBe('辞書にない単語: recieve');
  });

  it('enables Undo Last after an Ignore and steps back to the word', async () => {
    await open([
      [0, 0, 'teh'],
      [1, 0, 'recieve'],
    ]);
    button('無視').click();
    expect(disabled()).not.toContain('直前の修正を元に戻す');
    expect(button('キャンセル')).toBeTruthy();
    button('直前の修正を元に戻す').click();
    expect(labels()[0]).toBe('辞書にない単語: teh');
    expect(disabled()).toContain('直前の修正を元に戻す');
  });

  it('offers Delete for a repeated word and removes only the word', async () => {
    await open([
      [0, 0, 'the the cat'],
      [1, 0, 'teh'],
    ]);
    expect(labels()[0]).toBe('同じ語の繰り返し: the');
    expect(dialog()?.querySelector<HTMLInputElement>('#fc-spelling-change-to')?.value).toBe('');
    expect(dialog()?.querySelectorAll('.fc-spelling__suggestions [role="option"]')).toHaveLength(0);
    expect(disabled()).toEqual(['すべて無視', 'すべて削除', '候補表示', '直前の修正を元に戻す']);
    button('削除').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'the  cat' });
    expect(labels()[0]).toBe('辞書にない単語: teh');
    expect(button('閉じる')).toBeTruthy();
  });
});
