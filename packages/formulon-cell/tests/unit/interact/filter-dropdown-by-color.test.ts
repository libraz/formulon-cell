import { afterEach, describe, expect, it } from 'vitest';
import { NO_COLOR } from '../../../src/commands/filter.js';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import { en } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const RANGE = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

const seeded = (withFont = true): SpreadsheetStore => {
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    const formats = new Map(state.format.formats);
    (['N', 0, 50, 100] as const).forEach((value, row) => {
      cells.set(addrKey({ sheet: 0, row, col: 0 }), {
        value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
        formula: null,
      });
    });
    formats.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
      fill: '#FF0000',
      ...(withFont ? { color: '#0000ff' } : {}),
    });
    formats.set(addrKey({ sheet: 0, row: 3, col: 0 }), { fill: '#00ff00' });
    return {
      ...state,
      data: { ...state.data, cells },
      format: { ...state.format, formats },
      conditional: {
        ...state.conditional,
        rules: [
          {
            kind: 'icon-set',
            range: { sheet: 0, r0: 1, c0: 0, r1: 3, c1: 0 },
            icons: 'arrows3',
            showValue: true,
          },
        ],
      },
    };
  });
  return store;
};

const open = (store: SpreadsheetStore): void => {
  const host = document.createElement('div');
  host.className = 'fc-host';
  host.dataset.fcPlatform = 'mac';
  document.body.appendChild(host);
  attachFilterDropdown({
    store,
    host,
    strings: en,
    locale: 'en-US',
    history: new History(),
  }).open(RANGE, 0, { x: 10, y: 20, h: 24 });
};

const filterSide = () =>
  document.querySelector<HTMLElement>(
    '.fc-filter-dropdown [data-fc-color-side="filter"]',
  ) as HTMLElement;
const kind = () =>
  filterSide().querySelector<HTMLSelectElement>(
    '.fc-filter-dropdown__color-kind',
  ) as HTMLSelectElement;
const swatches = () => [
  ...filterSide().querySelectorAll<HTMLButtonElement>('.fc-filter-dropdown__swatch'),
];
const choose = (value: string): void => {
  kind().value = value;
  kind().dispatchEvent(new Event('change'));
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('Mac column filter by color', () => {
  it('offers every kind, an empty Cell Icon list included', () => {
    const store = seeded(false);
    store.setState((state) => ({
      ...state,
      conditional: { ...state.conditional, rules: [] },
    }));
    open(store);
    expect([...kind().options].map((o) => [o.value, o.disabled])).toEqual([
      ['', false],
      ['cellColor', false],
      ['fontColor', false],
      ['cellIcon', false],
    ]);
    choose('cellIcon');
    expect(swatches()).toEqual([]);
  });

  it('lists the column fills in first-appearance order, No Fill last, and filters by one', () => {
    const store = seeded();
    open(store);
    choose('cellColor');
    expect(swatches().map((b) => b.dataset.fcColor)).toEqual(['#ff0000', '#00ff00', NO_COLOR]);
    expect(swatches().at(-1)?.textContent).toBe(en.mac.filterPanel.noFill);
    swatches()[1]?.click();
    expect(store.getState().ui.filterCriteria[0]?.color).toEqual({
      kind: 'cellColor',
      color: '#00ff00',
    });
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([1, 2]);
    expect(kind().value).toBe('cellColor');
    expect(swatches().map((b) => b.getAttribute('aria-pressed'))).toEqual([
      'false',
      'true',
      'false',
    ]);
  });

  it('filters by No Fill and by an Automatic font colour', () => {
    const store = seeded();
    open(store);
    choose('cellColor');
    swatches().at(-1)?.click();
    expect(store.getState().ui.filterCriteria[0]?.color).toEqual({
      kind: 'cellColor',
      color: NO_COLOR,
    });
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([1, 3]);
    choose('fontColor');
    expect(swatches().map((b) => b.textContent || b.dataset.fcColor)).toEqual([
      '#0000ff',
      en.mac.filterPanel.fontAutomatic,
    ]);
    swatches().at(-1)?.click();
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([1]);
  });

  it('ends the sort side with Custom Sort..., which opens the Sort dialog from the header', () => {
    const store = seeded();
    const host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    let opened = 0;
    attachFilterDropdown({
      store,
      host,
      strings: en,
      locale: 'en-US',
      history: new History(),
      getWb: () => ({}) as never,
      onCustomSort: () => {
        opened += 1;
      },
    }).open(RANGE, 0, { x: 10, y: 20, h: 24 });
    const sortSide = document.querySelector<HTMLElement>(
      '.fc-filter-dropdown [data-fc-color-side="sort"]',
    ) as HTMLElement;
    const sortKind = sortSide.querySelector<HTMLSelectElement>(
      '.fc-filter-dropdown__color-kind',
    ) as HTMLSelectElement;
    sortKind.value = 'cellIcon';
    sortKind.dispatchEvent(new Event('change'));
    const last = sortSide.querySelector('.fc-filter-dropdown__swatches')?.lastElementChild;
    expect(last?.textContent).toBe(en.mac.filterPanel.customSort);
    (last as HTMLButtonElement).click();
    expect(opened).toBe(1);
    expect(document.querySelector('.fc-filter-dropdown')).toBeNull();
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('lists the icons the conditional format draws and filters by one', () => {
    const store = seeded();
    open(store);
    choose('cellIcon');
    expect(swatches().map((b) => b.dataset.fcIcon)).toEqual([
      'arrows3:0',
      'arrows3:1',
      'arrows3:2',
    ]);
    swatches()[0]?.click();
    expect(store.getState().ui.filterCriteria[0]?.color).toEqual({
      kind: 'cellIcon',
      icon: { set: 'arrows3', slot: 0 },
    });
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([2, 3]);
  });

  it('clears a colour filter when None is chosen', () => {
    const store = seeded();
    open(store);
    choose('fontColor');
    swatches()[0]?.click();
    expect(store.getState().layout.hiddenRows.size).toBe(2);
    choose('');
    expect(store.getState().ui.filterCriteria[0]?.color).toBeUndefined();
    expect(store.getState().layout.hiddenRows.size).toBe(0);
  });
});
