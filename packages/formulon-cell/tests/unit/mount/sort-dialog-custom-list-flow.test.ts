import { afterEach, describe, expect, it } from 'vitest';
import { customListsSourceFor } from '../../../src/commands/custom-lists.js';
import { createEditingDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/editing.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));
const primary = (root: Element | null | undefined) =>
  root?.closest('[role="dialog"]')?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');

describe('Sort dialog custom list', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('sorts by a list entered in Custom Lists and keeps it as a user list', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { instance } = sheet;
    ['Size', 'High', 'Low', 'Mid'].forEach((value, row) => {
      instance.workbook.setText({ sheet: 0, row, col: 0 }, value);
    });
    mutators.replaceCells(instance.store, instance.workbook.cells(0));
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 0 });
    // An existing list holding the same entries the other way round: the level must follow the
    // list it was given, not whichever list the entries also appear in.
    customListsSourceFor(instance.store)?.set([['High', 'Mid', 'Low']]);
    const done = createEditingDropdownDefaults(instance).applySortMenuAction('custom');
    await settle();
    const direction = document.querySelector<HTMLSelectElement>(
      '.fc-sortdlg .fc-sortdlg__direction',
    );
    if (!direction) throw new Error('missing order');
    direction.value = 'custom';
    direction.dispatchEvent(new Event('change'));
    await settle();
    const entries = document.querySelector<HTMLTextAreaElement>('.fc-customlists__entries');
    if (!entries) throw new Error('missing Custom Lists');
    entries.value = 'Low\nMid\nHigh';
    primary(entries)?.click();
    await settle();
    expect(direction.value).toBe('list');
    primary(document.querySelector('.fc-sortdlg'))?.click();
    await done;
    await settle();
    expect([1, 2, 3].map((row) => instance.workbook.getValue({ sheet: 0, row, col: 0 }))).toEqual([
      { kind: 'text', value: 'Low' },
      { kind: 'text', value: 'Mid' },
      { kind: 'text', value: 'High' },
    ]);
    expect(customListsSourceFor(instance.store)?.get()).toEqual([
      ['High', 'Mid', 'Low'],
      ['Low', 'Mid', 'High'],
    ]);
  });
});
