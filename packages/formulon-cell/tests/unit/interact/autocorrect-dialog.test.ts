import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { AutoCorrectOptions } from '../../../src/commands/autocorrect.js';
import { openAutoCorrectDialog } from '../../../src/interact/autocorrect-dialog.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const dialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-autocorrect:not([hidden])');
const button = (label: string): HTMLButtonElement => {
  const found = [...(dialog()?.querySelectorAll<HTMLButtonElement>('button') ?? [])].find(
    (candidate) => candidate.textContent === label,
  );
  if (!found) throw new Error(`no ${label} button`);
  return found;
};
const input = (id: string): HTMLInputElement =>
  dialog()?.querySelector<HTMLInputElement>(`#${id}`) as HTMLInputElement;
const type = (id: string, value: string): void => {
  input(id).value = value;
  input(id).dispatchEvent(new Event('input', { bubbles: true }));
};
const checkbox = (label: string): HTMLInputElement => {
  const row = [...(dialog()?.querySelectorAll('label') ?? [])].find(
    (candidate) => candidate.textContent === label,
  );
  return row?.querySelector('input') as HTMLInputElement;
};
const entries = (): string[] =>
  [...(dialog()?.querySelectorAll('[role="option"]') ?? [])].map((option) =>
    [...option.children].map((cell) => cell.textContent).join('→'),
  );

describe('AutoCorrect options dialog', () => {
  let sheet: MountedStubSheet;
  let options: AutoCorrectOptions = {};
  const setAutoCorrectOptions = vi.fn((next: AutoCorrectOptions) => {
    options = next;
  });

  beforeEach(async () => {
    options = {};
    setAutoCorrectOptions.mockClear();
    sheet = await mountStubSheet({
      locale: 'en',
      getAutoCorrectOptions: () => options,
      setAutoCorrectOptions,
    });
  });

  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  it('lists the replacements and labels itself as the desktop pane', () => {
    openAutoCorrectDialog(sheet.instance);
    expect(dialog()?.querySelector('.fc-fmtdlg__header')?.textContent).toBe('AutoCorrect Options');
    expect(entries()).toEqual(['(c)→©', '(e)→€', '(r)→®', '(tm)→™', '...→…']);
    expect(checkbox('Replace text as you type').checked).toBe(true);
    expect(checkbox('Replace internet and network paths with hyperlinks').checked).toBe(true);
    expect(button('Add').disabled).toBe(true);
    expect(button('Delete').disabled).toBe(true);
  });

  it('adds, replaces and deletes entries, committing each edit', () => {
    openAutoCorrectDialog(sheet.instance);
    type('fc-autocorrect-replace', 'teh');
    type('fc-autocorrect-with', 'the');
    button('Add').click();
    expect(entries()).toContain('teh→the');
    expect(setAutoCorrectOptions).toHaveBeenLastCalledWith({
      entries: { en: expect.arrayContaining([{ from: 'teh', to: 'the' }]) },
    });
    type('fc-autocorrect-replace', '(c)');
    expect(input('fc-autocorrect-with').value).toBe('©');
    button('Delete').click();
    expect(entries()).not.toContain('(c)→©');
    expect(sheet.instance.autoCorrect.correct('(c)')).toBe('(c)');
  });

  it('switches each option, and closes', () => {
    openAutoCorrectDialog(sheet.instance);
    checkbox('Replace text as you type').click();
    checkbox('Correct TWo INitial CApitals').click();
    checkbox('Capitalize first letter of sentences').click();
    checkbox('Capitalize names of days').click();
    checkbox('Replace internet and network paths with hyperlinks').click();
    expect(sheet.instance.autoCorrect.state()).toMatchObject({
      replaceAsYouType: false,
      replaceHyperlinks: false,
      correctTwoInitialCapitals: false,
      capitalizeSentences: false,
      capitalizeDays: false,
    });
    expect(setAutoCorrectOptions).toHaveBeenLastCalledWith(
      expect.objectContaining({ replaceAsYouType: false, replaceHyperlinks: false }),
    );
    button('Close').click();
    expect(dialog()).toBeNull();
  });
});
