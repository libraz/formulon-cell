import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachStatusBar } from '../../../src/interact/status-bar.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
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
    expect(statusbar.querySelector('.fc-host__statusbar-left')?.textContent).toContain('準備完了');
    expect(right?.textContent).toContain('セル');
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

  it('hides the calc-mode badge when getCalcMode is omitted or returns null', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      getCalcMode: () => null,
    });
    const badge = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-calcmode');
    expect(badge).not.toBeNull();
    expect(badge?.style.display).toBe('none');
    handle.detach();
  });

  it('renders the calc-mode badge with the active mode label', () => {
    let mode: 0 | 1 | 2 = 1; // Manual
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      getCalcMode: () => mode,
    });
    const badge = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-calcmode');
    expect(badge?.style.display).toBe('');
    // defaultStrings is ja-JP; the test asserts the localized label.
    expect(badge?.textContent).toContain('手動');
    expect(badge?.dataset.calcMode).toBe('1');

    mode = 0;
    handle.refresh();
    expect(badge?.textContent).toContain('自動');
    expect(badge?.dataset.calcMode).toBe('0');
    handle.detach();
  });

  it('badge click invokes onCycleCalcMode; double-click invokes onRecalc', () => {
    const cycle: number[] = [];
    const recalcs: number[] = [];
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      getCalcMode: () => 0,
      onCycleCalcMode: () => cycle.push(1),
      onRecalc: () => recalcs.push(1),
    });
    const badge = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-calcmode');
    badge?.click();
    badge?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));
    expect(cycle.length).toBe(1);
    expect(recalcs.length).toBe(1);
    handle.detach();
  });

  it('renders zoom controls and applies slider changes', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const slider = statusbar.querySelector<HTMLInputElement>('.fc-host__statusbar-zoom-slider');
    const label = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-zoom-label');
    expect(slider).not.toBeNull();
    expect(slider?.value).toBe('100');
    expect(label?.textContent).toBe('100%');
    expect(slider?.getAttribute('aria-label')).toBe('ズーム');

    if (!slider) throw new Error('expected zoom slider');
    slider.value = '150';
    slider.dispatchEvent(new Event('input'));
    expect(store.getState().viewport.zoom).toBe(1.5);
    expect(label?.textContent).toBe('150%');
    handle.detach();
  });

  it('setStrings relabels static status and zoom chrome', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    handle.setStrings(en);
    expect(statusbar.querySelector('.fc-host__statusbar-left')?.textContent).toContain('Ready');
    expect(statusbar.querySelector('.fc-host__statusbar-right')?.textContent).toContain('cell');
    expect(
      statusbar
        .querySelector<HTMLInputElement>('.fc-host__statusbar-zoom-slider')
        ?.getAttribute('aria-label'),
    ).toBe('Zoom');
    handle.detach();
  });

  it('delegates zoom changes when onZoomChange is provided', () => {
    const calls: number[] = [];
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      onZoomChange: (z) => {
        calls.push(z);
        mutators.setZoom(store, z);
      },
    });
    const plus = statusbar.querySelector<HTMLButtonElement>(
      '.fc-host__statusbar-zoom-btn:last-of-type',
    );
    plus?.click();
    expect(calls).toEqual([1.1]);
    expect(store.getState().viewport.zoom).toBeCloseTo(1.1);
    handle.detach();
  });
});
