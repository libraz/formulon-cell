import { afterEach, describe, expect, it } from 'vitest';
import { applyValueFilter } from '../../../src/commands/filter.js';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const RANGE = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

afterEach(() => {
  document.body.replaceChildren();
});

describe('Mac column panel sort by color', () => {
  it('puts the rows with the picked colour or icon first and shows it picked', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setText({ sheet: 0, row: 0, col: 0 }, 'N');
      [0, 50, 100].forEach((v, k) => {
        wb.setNumber({ sheet: 0, row: k + 1, col: 0 }, v);
      });
      const store = createSpreadsheetStore();
      mutators.replaceCells(store, wb.cells(0));
      store.setState((state) => ({
        ...state,
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
      }));
      applyValueFilter(store.getState(), store, RANGE, 0, [], wb);
      const host = document.createElement('div');
      host.className = 'fc-host';
      host.dataset.fcPlatform = 'mac';
      document.body.appendChild(host);
      attachFilterDropdown({
        store,
        host,
        strings: en,
        locale: 'en-US',
        getWb: () => wb,
        history: new History(),
      }).open(RANGE, 0, { x: 10, y: 20, h: 24 });
      const sortSide = () =>
        document.querySelector<HTMLElement>('[data-fc-color-side="sort"]') as HTMLElement;
      const kind = sortSide().querySelector<HTMLSelectElement>('select') as HTMLSelectElement;
      expect(kind.disabled).toBe(false);
      kind.value = 'cellIcon';
      kind.dispatchEvent(new Event('change'));
      sortSide()
        .querySelector<HTMLButtonElement>('.fc-filter-dropdown__swatch[data-fc-icon="arrows3:2"]')
        ?.click();
      expect([1, 2, 3].map((row) => wb.getValue({ sheet: 0, row, col: 0 }))).toEqual([
        { kind: 'number', value: 100 },
        { kind: 'number', value: 0 },
        { kind: 'number', value: 50 },
      ]);
      expect(store.getState().ui.filterSorts?.[0]?.keys).toEqual([
        { byCol: 0, direction: 'asc', sortOn: 'cellIcon', icon: { set: 'arrows3', slot: 2 } },
      ]);
      expect(
        sortSide()
          .querySelector('.fc-filter-dropdown__swatch[data-fc-icon="arrows3:2"]')
          ?.getAttribute('aria-pressed'),
      ).toBe('true');
    } finally {
      wb.dispose();
    }
  });
});
