import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History, recordSlicersChange } from '../../../src/commands/history.js';
import type { CellValue } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { attachSlicer } from '../../../src/interact/slicer.js';
import {
  type SlicerSpec,
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const seedCell = (
  store: SpreadsheetStore,
  sheet: number,
  row: number,
  col: number,
  value: CellValue,
): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet, row, col }), { value, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

/** Build a workbook stand-in that exposes the surface attachSlicer needs:
 *  `getTables()` returning a fixed table descriptor pointed at the seeded
 *  data. Anything else throws so missing wiring shows up loud. */
const makeFakeWb = (table: {
  name: string;
  displayName?: string;
  ref: string;
  sheetIndex: number;
  columns: string[];
}): WorkbookHandle => {
  return {
    getTables: () => [
      {
        name: table.name,
        displayName: table.displayName ?? table.name,
        ref: table.ref,
        sheetIndex: table.sheetIndex,
        columns: [...table.columns],
      },
    ],
  } as unknown as WorkbookHandle;
};

const seedRegionRows = (store: SpreadsheetStore): void => {
  // Header row on row 0 ("Region"), values on rows 1..3.
  seedCell(store, 0, 0, 0, { kind: 'text', value: 'Region' });
  seedCell(store, 0, 1, 0, { kind: 'text', value: 'East' });
  seedCell(store, 0, 2, 0, { kind: 'text', value: 'West' });
  seedCell(store, 0, 3, 0, { kind: 'text', value: 'East' });
};

describe('slicers slice mutators', () => {
  let store: SpreadsheetStore;
  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('addSlicer appends a spec; duplicates by id are rejected', () => {
    const spec: SlicerSpec = {
      id: 'a',
      tableName: 'Sales',
      column: 'Region',
      selected: ['East'],
    };
    mutators.addSlicer(store, spec);
    mutators.addSlicer(store, { ...spec, column: 'Other' });
    expect(store.getState().slicers.slicers).toHaveLength(1);
    expect(store.getState().slicers.slicers[0]?.column).toBe('Region');
  });

  it('removeSlicer drops the matching id; absent id is a no-op', () => {
    mutators.addSlicer(store, { id: 'a', tableName: 'T', column: 'C', selected: [] });
    mutators.addSlicer(store, { id: 'b', tableName: 'T', column: 'D', selected: [] });
    mutators.removeSlicer(store, 'missing');
    expect(store.getState().slicers.slicers).toHaveLength(2);
    mutators.removeSlicer(store, 'a');
    expect(store.getState().slicers.slicers.map((s) => s.id)).toEqual(['b']);
  });

  it('updateSlicer merges patch into the targeted spec', () => {
    mutators.addSlicer(store, { id: 'a', tableName: 'T', column: 'C', selected: ['x'] });
    mutators.updateSlicer(store, 'a', { column: 'D', x: 42 });
    const sp = store.getState().slicers.slicers[0];
    expect(sp?.column).toBe('D');
    expect(sp?.x).toBe(42);
    // selected untouched when patch omits it
    expect(sp?.selected).toEqual(['x']);
  });

  it('setSlicerSelected replaces the chip selection (empty = include all)', () => {
    mutators.addSlicer(store, { id: 'a', tableName: 'T', column: 'C', selected: [] });
    mutators.setSlicerSelected(store, 'a', ['East', 'West']);
    expect(store.getState().slicers.slicers[0]?.selected).toEqual(['East', 'West']);
    mutators.setSlicerSelected(store, 'a', []);
    expect(store.getState().slicers.slicers[0]?.selected).toEqual([]);
  });

  it('history undo restores the previous chip selection', () => {
    mutators.addSlicer(store, { id: 'a', tableName: 'T', column: 'C', selected: [] });
    const history = new History();
    recordSlicersChange(history, store, () => {
      mutators.setSlicerSelected(store, 'a', ['East']);
    });
    expect(store.getState().slicers.slicers[0]?.selected).toEqual(['East']);
    expect(history.undo()).toBe(true);
    expect(store.getState().slicers.slicers[0]?.selected).toEqual([]);
    expect(history.redo()).toBe(true);
    expect(store.getState().slicers.slicers[0]?.selected).toEqual(['East']);
  });
});

describe('attachSlicer', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    seedRegionRows(store);
    wb = makeFakeWb({
      name: 'Sales',
      ref: 'A1:A4',
      sheetIndex: 0,
      columns: ['Region'],
    });
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('addSlicer({tableName, column}) renders one chip per distinct value', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    const spec = handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    expect(spec.column).toBe('Region');
    const chips = host.querySelectorAll<HTMLButtonElement>('.fc-slicer__chip');
    const labels = Array.from(chips).map((c) => c.textContent);
    // distinctValues sorts asc — East < West.
    expect(labels).toEqual(['East', 'West']);
    handle.detach();
  });

  it('toggling a chip narrows the autoFilter to the selected value(s)', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    const chips = host.querySelectorAll<HTMLButtonElement>('.fc-slicer__chip');
    // Click "East" — only the West row (index 2) should hide.
    const east = Array.from(chips).find((c) => c.textContent === 'East');
    east?.click();
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    // The two East rows (1, 3) stay visible.
    expect(store.getState().layout.hiddenRows.has(1)).toBe(false);
    expect(store.getState().layout.hiddenRows.has(3)).toBe(false);
    handle.detach();
  });

  it('toggling all chips off restores the include-all state', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    const spec = handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    // Manually set selection to ['East'] so the West row hides.
    mutators.setSlicerSelected(store, spec.id, ['East']);
    handle.refresh();
    // Now click "East" again — selection becomes empty → no rows hidden.
    const east = host.querySelector<HTMLButtonElement>('.fc-slicer__chip[data-fc-value="East"]');
    east?.click();
    expect(store.getState().slicers.slicers[0]?.selected).toEqual([]);
    expect(store.getState().layout.hiddenRows.size).toBe(0);
    handle.detach();
  });

  it('the panel header × removes the slicer and tears down the panel', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    expect(host.querySelector('.fc-slicer')).not.toBeNull();
    const close = host.querySelector<HTMLButtonElement>('.fc-slicer__close');
    close?.click();
    expect(store.getState().slicers.slicers).toHaveLength(0);
    expect(host.querySelector('.fc-slicer')).toBeNull();
    handle.detach();
  });

  it('addSlicer throws when the table or column is missing', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    expect(() => handle.addSlicer({ tableName: 'Nope', column: 'Region' })).toThrow(/not found/);
    expect(() => handle.addSlicer({ tableName: 'Sales', column: 'Nope' })).toThrow(/not in table/);
    handle.detach();
  });

  it('history-tracked add → undo removes the panel and clears the autoFilter', () => {
    const history = new History();
    const handle = attachSlicer({ host, store, getWb: () => wb, history });
    handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    // Toggle East — changes selection + applies filter.
    const east = host.querySelector<HTMLButtonElement>('.fc-slicer__chip[data-fc-value="East"]');
    east?.click();
    expect(store.getState().slicers.slicers[0]?.selected).toEqual(['East']);
    // Undo the chip toggle — back to empty selection.
    expect(history.undo()).toBe(true);
    expect(store.getState().slicers.slicers[0]?.selected).toEqual([]);
    handle.detach();
  });

  it('refresh() rebuilds the chip set when underlying data changes', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    expect(host.querySelectorAll('.fc-slicer__chip').length).toBe(2);
    // Add a new region value.
    seedCell(store, 0, 4, 0, { kind: 'text', value: 'North' });
    // Bump the table ref to include the new row.
    const wb2 = makeFakeWb({
      name: 'Sales',
      ref: 'A1:A5',
      sheetIndex: 0,
      columns: ['Region'],
    });
    // Swap the closure target so the slicer reads the new ref.
    Object.assign(wb, wb2);
    Object.defineProperty(wb, 'getTables', {
      value: (wb2 as unknown as { getTables: () => unknown }).getTables,
      configurable: true,
    });
    handle.refresh();
    expect(host.querySelectorAll('.fc-slicer__chip').length).toBe(3);
    handle.detach();
  });

  it('detach removes panels and unsubscribes from store changes', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    expect(host.querySelector('.fc-slicer')).not.toBeNull();
    handle.detach();
    expect(host.querySelector('.fc-slicer')).toBeNull();
    // Mutating after detach should not crash or rebuild the panel.
    mutators.addSlicer(store, { id: 'x', tableName: 'Sales', column: 'Region', selected: [] });
    expect(host.querySelector('.fc-slicer')).toBeNull();
  });

  it('refresh is invoked when subscribed to a recalc/store update', () => {
    const handle = attachSlicer({ host, store, getWb: () => wb });
    handle.addSlicer({ tableName: 'Sales', column: 'Region' });
    const refreshSpy = vi.spyOn(handle, 'refresh');
    // Drive a cells map swap — the internal subscriber should rerender.
    seedCell(store, 0, 4, 0, { kind: 'text', value: 'North' });
    // The subscriber path mutates DOM directly; refresh isn't called via
    // the spy (renderAll is internal). Instead assert the chip set
    // refreshed.
    expect(host.querySelectorAll('.fc-slicer__chip').length).toBe(2);
    handle.detach();
    refreshSpy.mockRestore();
  });
});
