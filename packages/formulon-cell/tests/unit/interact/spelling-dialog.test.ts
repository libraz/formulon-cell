import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setComment } from '../../../src/commands/comment.js';
import { createDefaultRibbonHooks } from '../../../src/mount/toolbar-defaults.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { seedText } from '../mount/toolbar/fixtures.js';

const dialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-spelling:not([hidden])');
const button = (label: string): HTMLButtonElement => {
  const found = [...(dialog()?.querySelectorAll<HTMLButtonElement>('button') ?? [])].find(
    (candidate) => candidate.textContent === label,
  );
  if (!found) throw new Error(`no ${label} button`);
  return found;
};
const flagged = (): string =>
  dialog()?.querySelector<HTMLInputElement>('#fc-spelling-word')?.value ?? '';
const changeTo = (): HTMLInputElement =>
  dialog()?.querySelector<HTMLInputElement>('#fc-spelling-change-to') as HTMLInputElement;

describe('Review > Spelling dialog', () => {
  let sheet: MountedStubSheet;

  const text = (row: number, col: number): unknown =>
    sheet.workbook.getValue({ sheet: 0, row, col });
  const select = (active: [number, number], range: [number, number, number, number]): void => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: active[0], col: active[1] });
    const [r0, c0, r1, c1] = range;
    mutators.setRange(sheet.instance.store, { sheet: 0, r0, c0, r1, c1 });
  };
  const spell = (): void => {
    createDefaultRibbonHooks(sheet.instance).review?.spelling?.();
  };
  const activeLabel = (): string => {
    const { row, col } = sheet.instance.store.getState().selection.active;
    return `${String.fromCharCode(65 + col)}${row + 1}`;
  };

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    // The Book18 layout: typos at A1 and C5, a clean word at B2.
    seedText(sheet, 0, 0, 'teh');
    seedText(sheet, 1, 1, 'hello');
    seedText(sheet, 4, 2, 'recieve');
  });

  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  it('walks the whole sheet by rows from one selected cell, wraps silently and closes after the last', () => {
    select([1, 1], [1, 1, 1, 1]);
    spell();
    expect(flagged()).toBe('recieve');
    expect(activeLabel()).toBe('C5');
    button('Ignore Once').click();
    expect(flagged()).toBe('teh');
    expect(activeLabel()).toBe('A1');
    button('Ignore Once').click();
    expect(dialog()).toBeNull();
  });

  it('checks only a selected range', () => {
    select([0, 0], [0, 0, 2, 0]);
    spell();
    expect(flagged()).toBe('teh');
    expect(sheet.instance.store.getState().selection.range).toMatchObject({ r0: 0, r1: 2 });
    button('Ignore Once').click();
    expect(dialog()).toBeNull();
  });

  it('opens nothing when nothing is flagged', () => {
    select([1, 1], [1, 1, 1, 1]);
    seedText(sheet, 0, 0, 'the');
    seedText(sheet, 4, 2, 'receive');
    spell();
    expect(dialog()).toBeNull();
  });

  it('labels itself as the desktop dialog', () => {
    select([0, 0], [0, 0, 0, 0]);
    spell();
    expect(dialog()?.querySelector('.fc-fmtdlg__header')?.textContent).toBe('Spelling: English');
    expect(dialog()?.textContent).toContain('Not in Dictionary:');
    expect(changeTo().value).toBe('the');
    expect(button('Cancel')).toBeTruthy();
  });

  it('adds the word and its correction to AutoCorrect, then changes it', () => {
    select([0, 0], [0, 0, 0, 0]);
    spell();
    button('AutoCorrect').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'the' });
    expect(sheet.instance.autoCorrect.state().entries).toContainEqual({ from: 'teh', to: 'the' });
    expect(flagged()).toBe('recieve');
  });

  it('changes the word as one undo step, then offers Close and Undo Last', () => {
    select([0, 0], [0, 0, 0, 0]);
    spell();
    button('Change').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'the' });
    expect(flagged()).toBe('recieve');
    expect(button('Close')).toBeTruthy();
    button('Undo Last').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'teh' });
    expect(flagged()).toBe('teh');
  });

  it('changes every occurrence in one undo step with Change All', () => {
    seedText(sheet, 2, 0, 'teh and teh');
    select([0, 0], [0, 0, 0, 0]);
    spell();
    button('Change All').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'the' });
    expect(text(2, 0)).toEqual({ kind: 'text', value: 'the and the' });
    expect(flagged()).toBe('recieve');
    button('Close').click();
    expect(sheet.instance.history.undo()).toBe(true);
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'teh' });
    expect(text(2, 0)).toEqual({ kind: 'text', value: 'teh and teh' });
  });

  it('keeps Ignore All and Add for the session, and hands added words to the host', async () => {
    const addWord = vi.fn();
    sheet.dispose();
    sheet = await mountStubSheet({
      locale: 'en',
      getSpellingService: () => ({
        check: (word) => word !== 'zork' && word !== 'blarg',
        suggest: () => ['fork'],
        addWord,
      }),
    });
    seedText(sheet, 0, 0, 'zork blarg');
    seedText(sheet, 1, 0, 'zork blarg');
    select([0, 0], [0, 0, 0, 0]);
    spell();
    expect(flagged()).toBe('zork');
    expect(changeTo().value).toBe('fork');
    button('Ignore All').click();
    expect(flagged()).toBe('blarg');
    button('Add').click();
    expect(addWord).toHaveBeenCalledWith('blarg', 'en');
    expect(dialog()).toBeNull();
    spell();
    expect(dialog()).toBeNull();
  });

  it('deletes a repeated word', () => {
    seedText(sheet, 0, 0, 'the the end');
    select([0, 0], [0, 0, 0, 0]);
    spell();
    expect(dialog()?.textContent).toContain('Repeated Word:');
    expect(changeTo().value).toBe('');
    button('Delete').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'the  end' });
  });

  it('checks a note after its cell and changes it as one undo step', () => {
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'see teh note', sheet.workbook);
    select([0, 0], [0, 0, 0, 0]);
    spell();
    button('Ignore Once').click();
    expect(flagged()).toBe('teh');
    button('Change').click();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.comment).toBe(
      'see the note',
    );
    sheet.instance.history.undo();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.comment).toBe(
      'see teh note',
    );
  });

  it('refuses a change the policy denies and stays on the word', () => {
    sheet.instance.commands.setPolicy({
      defaultOperation: 'allow',
      operations: { valueEdit: false },
    });
    select([0, 0], [0, 0, 0, 0]);
    spell();
    button('Change').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'teh' });
    expect(flagged()).toBe('teh');
    expect(dialog()?.querySelector('.fc-spelling__status')?.getAttribute('data-state')).toBe(
      'error',
    );
  });
});
