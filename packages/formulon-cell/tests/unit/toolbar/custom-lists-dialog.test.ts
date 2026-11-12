import { afterEach, describe, expect, it, vi } from 'vitest';
import { showCustomListsDialog } from '../../../src/toolbar/dialogs/custom-lists.js';
import { showSortDialog } from '../../../src/toolbar/dialogs/sort.js';

const labels = {
  title: 'Custom Lists',
  newList: 'NEW LIST',
  lists: 'Custom lists:',
  entries: 'List entries:',
  add: 'Add',
  delete: 'Delete',
  hint: 'Press Return to separate list entries.',
  ok: 'OK',
  cancel: 'Cancel',
};
const BUILT_INS = [['Sun', 'Mon', 'Tue']];

const q = <T extends HTMLElement>(selector: string): T =>
  document.querySelector<T>(`.fc-customlists ${selector}`) as T;
const select = (value: string): void => {
  const lists = q<HTMLSelectElement>('.fc-customlists__lists');
  lists.value = value;
  lists.dispatchEvent(new Event('change'));
};
const ok = (): void =>
  document
    .querySelector('.fc-customlists')
    ?.closest('[role="dialog"]')
    ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
    ?.click();

afterEach(() => {
  document.body.replaceChildren();
});

describe('Custom Lists dialog', () => {
  it('lists NEW LIST, the built-in lists and the user lists, with built-ins read-only', () => {
    void showCustomListsDialog({
      labels,
      builtIns: BUILT_INS,
      userLists: [['Low', 'High']],
      onChange: () => undefined,
    });
    expect(
      [...q<HTMLSelectElement>('.fc-customlists__lists').options].map((o) => o.textContent),
    ).toEqual(['NEW LIST', 'Sun, Mon, Tue', 'Low, High']);
    select('b:0');
    expect(q<HTMLTextAreaElement>('.fc-customlists__entries').readOnly).toBe(true);
    expect(q<HTMLButtonElement>('.fc-customlists__delete').disabled).toBe(true);
    expect(q<HTMLButtonElement>('.fc-customlists__add').disabled).toBe(true);
    select('u:0');
    expect(q<HTMLTextAreaElement>('.fc-customlists__entries').value).toBe('Low\nHigh');
    expect(q<HTMLButtonElement>('.fc-customlists__delete').disabled).toBe(false);
  });

  it('adds a list typed under NEW LIST and deletes a user list, reporting each change', () => {
    const onChange = vi.fn();
    void showCustomListsDialog({ labels, builtIns: BUILT_INS, userLists: [], onChange });
    q<HTMLTextAreaElement>('.fc-customlists__entries').value = 'North\nSouth, East';
    q<HTMLButtonElement>('.fc-customlists__add').click();
    expect(onChange).toHaveBeenLastCalledWith([['North', 'South', 'East']]);
    expect(q<HTMLSelectElement>('.fc-customlists__lists').value).toBe('u:0');
    q<HTMLButtonElement>('.fc-customlists__delete').click();
    expect(onChange).toHaveBeenLastCalledWith([]);
  });

  it('resolves with the selected list on OK, adding a typed new list first', async () => {
    const onChange = vi.fn();
    const pending = showCustomListsDialog({ labels, builtIns: BUILT_INS, userLists: [], onChange });
    q<HTMLTextAreaElement>('.fc-customlists__entries').value = 'S\nM\nL';
    ok();
    await expect(pending).resolves.toEqual(['S', 'M', 'L']);
    expect(onChange).toHaveBeenCalledWith([['S', 'M', 'L']]);
  });

  it('keeps Return in the entries box for a new line instead of pressing OK', async () => {
    let settled = false;
    void showCustomListsDialog({
      labels,
      builtIns: BUILT_INS,
      userLists: [],
      onChange: () => undefined,
    }).then(() => {
      settled = true;
    });
    const entries = q<HTMLTextAreaElement>('.fc-customlists__entries');
    entries.focus();
    entries.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await Promise.resolve();
    expect(settled).toBe(false);
    expect(document.querySelector('.fc-customlists')).not.toBeNull();
  });

  it('opens on the list a level already uses', () => {
    void showCustomListsDialog({
      labels,
      builtIns: BUILT_INS,
      userLists: [['Low', 'High']],
      onChange: () => undefined,
      initial: ['Low', 'High'],
    });
    expect(q<HTMLSelectElement>('.fc-customlists__lists').value).toBe('u:0');
  });
});

describe('Sort dialog Order › Custom List…', () => {
  it('orders a level by the list the Custom Lists dialog returns', async () => {
    const open = vi.fn(async () => ['Low', 'Mid', 'High']);
    const pending = showSortDialog({
      title: 'Sort',
      columnLabel: 'Sort by',
      thenByLabel: 'Then by',
      noThenByLabel: '(none)',
      orderLabel: 'Order',
      headerLabel: 'My data has headers',
      addLevelLabel: 'Add Level',
      deleteLevelLabel: 'Delete Level',
      copyLevelLabel: 'Copy Level',
      levelUnavailableLabel: 'No level selected',
      ascendingLabel: 'A to Z',
      descendingLabel: 'Z to A',
      sortOn: {
        heading: 'Sort On',
        values: 'Cell Values',
        cellColor: 'Cell Color',
        fontColor: 'Font Color',
        cellIcon: 'Conditional Formatting Icon',
        onTop: 'On Top',
        onBottom: 'On Bottom',
      },
      customList: { label: 'Custom List...', open },
      columns: [{ value: '0', label: 'Column A' }],
      initialColumn: '0',
      initialDirection: 'asc',
      initialHasHeader: true,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
    });
    const direction = document.querySelector<HTMLSelectElement>('.fc-sortdlg__direction');
    if (!direction) throw new Error('missing order');
    expect([...direction.options].map((o) => o.textContent)).toEqual([
      'A to Z',
      'Z to A',
      'Custom List...',
    ]);
    direction.value = 'custom';
    direction.dispatchEvent(new Event('change'));
    await Promise.resolve();
    await Promise.resolve();
    expect(open).toHaveBeenCalledWith(null);
    expect(direction.value).toBe('list');
    expect(direction.selectedOptions[0]?.textContent).toBe('Low, Mid, High');
    document.querySelector<HTMLButtonElement>('[role="dialog"] .fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toMatchObject({
      levels: [
        { column: '0', direction: 'asc', sortOn: 'customList', customList: ['Low', 'Mid', 'High'] },
      ],
    });
  });
});
