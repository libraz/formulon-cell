import { FilterKind, IconSetName } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { applyColorFilter } from '../../../src/commands/filter.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { FilterColorSpec } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

const seeded = async () => {
  const wb = await WorkbookHandle.createDefault();
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    const formats = new Map(state.format.formats);
    const values: (string | number)[] = ['N', 0, 50, 100];
    values.forEach((value, row) => {
      const addr = { sheet: 0, row, col: 0 };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
      cells.set(addrKey(addr), {
        value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
        formula: null,
      });
    });
    formats.set(addrKey({ sheet: 0, row: 1, col: 0 }), { fill: '#ff0000', color: '#0000ff' });
    formats.set(addrKey({ sheet: 0, row: 3, col: 0 }), { fill: '#ff0000' });
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
  return { wb, store };
};

describe('AutoFilter by colour and by icon', () => {
  const cases: { color: FilterColorSpec; native: object; shown: number[] }[] = [
    {
      color: { kind: 'cellColor', color: '#ff0000' },
      native: { kind: FilterKind.Color, cellColor: true },
      shown: [1, 3],
    },
    {
      color: { kind: 'fontColor', color: '#0000ff' },
      native: { kind: FilterKind.Color, cellColor: false },
      shown: [1],
    },
    {
      color: { kind: 'cellIcon', icon: { set: 'arrows3', slot: 2 } },
      native: { kind: FilterKind.Icon, iconSet: IconSetName.ThreeArrows, iconIndex: 2 },
      shown: [3],
    },
  ];

  it('keeps the matching rows, writes the native column and reads it back after a save', async () => {
    for (const { color, native, shown } of cases) {
      const { wb, store } = await seeded();
      try {
        applyColorFilter(store.getState(), store, range, 0, color);
        const hidden = store.getState().layout.hiddenRows;
        expect(
          [1, 2, 3].filter((row) => !hidden.has(row)),
          color.kind,
        ).toEqual(shown);
        expect(syncAutoFilterToEngine(wb, store.getState(), 0, store), color.kind).toBe(true);
        expect(wb.getAutoFilter(0)?.columns[0], color.kind).toMatchObject(native);
        const reopened = await WorkbookHandle.loadBytes(wb.save());
        try {
          const reloaded = createSpreadsheetStore();
          hydrateAutoFilterFromEngine(reopened, reloaded, 0);
          expect(reloaded.getState().ui.filterCriteria[0]?.color, color.kind).toEqual(color);
        } finally {
          reopened.dispose();
        }
      } finally {
        wb.dispose();
      }
    }
  });
});
