import { afterEach, beforeEach, describe, expect, it } from 'vitest';
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
const status = (): HTMLElement | null => dialog()?.querySelector('.fc-spelling__status') ?? null;

describe('Review > Spelling dialog write failures', () => {
  let sheet: MountedStubSheet;
  const text = (row: number, col: number): unknown =>
    sheet.workbook.getValue({ sheet: 0, row, col });
  const spell = (): void => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    createDefaultRibbonHooks(sheet.instance).review?.spelling?.();
  };

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    seedText(sheet, 0, 0, 'teh');
    seedText(sheet, 1, 0, 'recieve');
  });
  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  it('applies nothing and shows a reason when Change All is refused on a note', () => {
    seedText(sheet, 2, 0, 'teh and teh');
    setComment(sheet.instance.store, { sheet: 0, row: 3, col: 0 }, 'a teh note', sheet.workbook);
    sheet.instance.commands.setPolicy({
      defaultOperation: 'allow',
      operations: { comment: false },
    });
    sheet.instance.history.clear();
    spell();
    button('Change All').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'teh' });
    expect(text(2, 0)).toEqual({ kind: 'text', value: 'teh and teh' });
    expect(sheet.instance.store.getState().format.formats.get('0:3:0')?.comment).toBe('a teh note');
    expect(flagged()).toBe('teh');
    expect(status()?.getAttribute('data-state')).toBe('error');
    expect(status()?.textContent).not.toBe('');
    expect(sheet.instance.history.canUndo()).toBe(false);
  });

  it('adds the AutoCorrect entry only after the change is written', () => {
    sheet.instance.commands.setPolicy({
      defaultOperation: 'allow',
      operations: { valueEdit: false },
    });
    spell();
    button('AutoCorrect').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'teh' });
    expect(sheet.instance.autoCorrect.state().entries).not.toContainEqual({
      from: 'teh',
      to: 'the',
    });
  });

  it('keeps the Undo Last record when the undo fails', () => {
    spell();
    button('Change').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'the' });
    const real = sheet.instance.history.undo.bind(sheet.instance.history);
    let fail = true;
    sheet.instance.history.undo = () => {
      if (fail) {
        fail = false;
        return false;
      }
      return real();
    };
    button('Undo Last').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'the' });
    button('Undo Last').click();
    expect(text(0, 0)).toEqual({ kind: 'text', value: 'teh' });
    expect(flagged()).toBe('teh');
  });
});
