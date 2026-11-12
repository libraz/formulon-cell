import { afterEach, describe, expect, it } from 'vitest';
import { showSortDialog } from '../../../src/toolbar/dialogs/sort.js';

const labels = {
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
  okLabel: 'OK',
  cancelLabel: 'Cancel',
};

const choices = {
  '0': { fills: ['#ff0000', '#00ff00'], fonts: [], icons: [] },
  '1': { fills: [], fonts: [], icons: [{ set: 'arrows3' as const, slot: 2 }] },
};

const open = () =>
  showSortDialog({
    ...labels,
    columns: [
      { value: '0', label: 'Column A' },
      { value: '1', label: 'Column B' },
    ],
    colorChoices: (column) => choices[column as '0' | '1'],
    initialColumn: '0',
    initialDirection: 'asc',
    initialHasHeader: true,
  });

const row = () => document.querySelector<HTMLElement>('.fc-sortdlg__level') as HTMLElement;
const pick = (select: HTMLSelectElement | null, value: string): void => {
  if (!select) throw new Error('missing select');
  select.value = value;
  select.dispatchEvent(new Event('change'));
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('Sort dialog Sort On', () => {
  it('offers the colour kinds the column shows and orders a colour On Top or On Bottom', async () => {
    const pending = open();
    const sortOn = row().querySelector<HTMLSelectElement>('.fc-sortdlg__sort-on');
    expect(Array.from(sortOn?.options ?? [], (o) => [o.value, o.disabled])).toEqual([
      ['values', false],
      ['cellColor', false],
      ['fontColor', true],
      ['cellIcon', true],
    ]);
    pick(sortOn, 'cellColor');
    const direction = row().querySelector<HTMLSelectElement>('.fc-sortdlg__direction');
    expect(Array.from(direction?.options ?? [], (o) => o.text)).toEqual(['On Top', 'On Bottom']);
    const swatches = [...row().querySelectorAll<HTMLButtonElement>('.fc-sortdlg__swatch')];
    expect(swatches.map((b) => [b.dataset.fcColor, b.getAttribute('aria-pressed')])).toEqual([
      ['#ff0000', 'true'],
      ['#00ff00', 'false'],
    ]);
    swatches[1]?.click();
    pick(direction, 'desc');
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toMatchObject({
      levels: [{ column: '0', direction: 'desc', sortOn: 'cellColor', color: '#00ff00' }],
    });
  });

  it('switches the kinds with the column and picks an icon level', async () => {
    const pending = open();
    pick(row().querySelector<HTMLSelectElement>('.fc-sortdlg__column'), '1');
    const sortOn = row().querySelector<HTMLSelectElement>('.fc-sortdlg__sort-on');
    expect(Array.from(sortOn?.options ?? [], (o) => o.disabled)).toEqual([
      false,
      true,
      true,
      false,
    ]);
    pick(sortOn, 'cellIcon');
    expect(
      [...row().querySelectorAll<HTMLButtonElement>('.fc-sortdlg__swatch')].map(
        (b) => b.dataset.fcIcon,
      ),
    ).toEqual(['arrows3:2']);
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toMatchObject({
      levels: [
        { column: '1', direction: 'asc', sortOn: 'cellIcon', icon: { set: 'arrows3', slot: 2 } },
      ],
    });
  });

  it('keeps a value level plain', async () => {
    const pending = open();
    expect(row().querySelector<HTMLElement>('.fc-sortdlg__color')?.hidden).toBe(true);
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    const result = await pending;
    expect(result?.levels).toEqual([{ column: '0', direction: 'asc' }]);
  });
});
