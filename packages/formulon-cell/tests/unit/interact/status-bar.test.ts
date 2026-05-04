import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { attachStatusBar } from '../../../src/interact/status-bar.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const seedNumber = (store: SpreadsheetStore, row: number, col: number, value: number): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('attachStatusBar', () => {
  let statusbar: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    statusbar = document.createElement('div');
    document.body.appendChild(statusbar);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders left/center/right segments and shows engine label on the right', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    expect(statusbar.querySelector('.fc-host__statusbar-left')).not.toBeNull();
    expect(statusbar.querySelector('.fc-host__statusbar-aggs')).not.toBeNull();
    const right = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-right');
    expect(right?.textContent).toContain('stub');
    handle.detach();
  });

  it('reflects sum/avg/count for a numeric selection', () => {
    seedNumber(store, 0, 0, 10);
    seedNumber(store, 1, 0, 20);
    seedNumber(store, 2, 0, 30);
    setRange(store, 0, 0, 2, 0);

    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const center = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-aggs');
    expect(center?.textContent).toContain('60'); // sum
    expect(center?.textContent).toContain('20'); // avg
    expect(center?.textContent).toContain('3'); // count
    handle.detach();
  });

  it('keeps center empty when nothing is selected (and no aggs apply)', () => {
    setRange(store, 0, 0, 0, 0);
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const center = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-aggs');
    expect(center?.textContent).toBe('');
    handle.detach();
  });

  it('right-click opens a chooser; clicking a row toggles the agg in the store', () => {
    seedNumber(store, 0, 0, 5);
    setRange(store, 0, 0, 0, 0);
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });

    statusbar.dispatchEvent(
      new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        clientX: 100,
        clientY: 200,
      }),
    );
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    expect(chooser).not.toBeNull();
    expect(chooser?.style.display).toBe('block');

    const items = chooser?.querySelectorAll<HTMLButtonElement>('.fc-statusbar__chooser-item');
    expect(items?.length).toBe(6);
    // Toggle "sum" off — initial set is ['sum','average','count'].
    const sumItem = Array.from(items ?? []).find((b) => b.textContent?.includes('合計'));
    expect(sumItem).toBeDefined();
    sumItem?.click();
    expect(store.getState().ui.statusAggs).not.toContain('sum');
    handle.detach();
  });

  it('Escape closes the chooser', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    statusbar.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    expect(chooser?.style.display).toBe('none');
    handle.detach();
  });

  it('refresh() re-reads the engine label', () => {
    let label = 'stub';
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => label,
    });
    let right = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-right');
    expect(right?.textContent).toContain('stub');
    label = 'formulon 9.9.9';
    handle.refresh();
    right = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-right');
    expect(right?.textContent).toContain('formulon 9.9.9');
    handle.detach();
  });

  it('detach removes the chooser node and stops subscribing', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    handle.detach();
    expect(document.querySelector('.fc-statusbar__chooser')).toBeNull();

    // Mutating state after detach should not crash.
    mutators.setActive(store, { sheet: 0, row: 5, col: 5 });
  });
});
