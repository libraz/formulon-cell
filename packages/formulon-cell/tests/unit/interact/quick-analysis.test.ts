import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachQuickAnalysis } from '../../../src/interact/quick-analysis.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seed = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

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
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('attachQuickAnalysis', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('mounts a hidden popover under the host on attach', () => {
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    const root = host.querySelector<HTMLElement>('.fc-quick');
    expect(root).not.toBeNull();
    expect(root?.hidden).toBe(true);
    handle.detach();
  });

  it('open() reveals the popover and renders a section per non-empty group', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 0, col: 2, value: 3 },
    ]);
    setRange(store, 0, 0, 0, 2);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();
    const root = host.querySelector<HTMLElement>('.fc-quick');
    expect(root?.hidden).toBe(false);
    const sections = root?.querySelectorAll('.fc-quick__section') ?? [];
    // Formatting + Totals + Tables + Sparklines + Charts (5 max). Empty groups are skipped.
    expect(sections.length).toBeGreaterThanOrEqual(3);
    handle.close();
    expect(root?.hidden).toBe(true);
    handle.detach();
  });

  it('groups surface category-specific actions (sparkline-line for a horizontal numeric run)', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 0, col: 2, value: 3 },
      { row: 0, col: 3, value: 4 },
    ]);
    setRange(store, 0, 0, 0, 3);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();
    const sparkBtn = host.querySelector<HTMLButtonElement>('button[data-action="sparkline-line"]');
    const totalSumBtn = host.querySelector<HTMLButtonElement>(
      'button[data-action="totals-sum-row"]',
    );
    const dataBarBtn = host.querySelector<HTMLButtonElement>(
      'button[data-action="format-data-bar"]',
    );
    expect(sparkBtn).not.toBeNull();
    expect(sparkBtn?.disabled).toBe(false);
    expect(totalSumBtn?.disabled).toBe(false);
    expect(dataBarBtn?.disabled).toBe(false);
    handle.detach();
  });

  it('clicking format-data-bar invokes executeQuickAnalysisAction → adds a conditional rule', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 5 },
      { row: 0, col: 1, value: 10 },
      { row: 0, col: 2, value: 15 },
    ]);
    setRange(store, 0, 0, 0, 2);
    expect(store.getState().conditional.rules).toHaveLength(0);

    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();
    const dataBarBtn = host.querySelector<HTMLButtonElement>(
      'button[data-action="format-data-bar"]',
    );
    dataBarBtn?.click();
    const rules = store.getState().conditional.rules;
    expect(rules).toHaveLength(1);
    expect(rules[0]?.kind).toBe('data-bar');
    handle.detach();
  });

  it('outside pointerdown closes an open popover', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();
    const root = host.querySelector<HTMLElement>('.fc-quick');
    expect(root?.hidden).toBe(false);
    // Click outside the popover but still inside the host — handler is on host.
    host.dispatchEvent(
      new PointerEvent('pointerdown', { bubbles: true, cancelable: true, pointerId: 1 }),
    );
    expect(root?.hidden).toBe(true);
    handle.detach();
  });

  it('Escape closes the popover and restores focus to the opener', () => {
    host.tabIndex = -1;
    seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    host.focus();
    handle.open();
    const root = host.querySelector<HTMLElement>('.fc-quick');
    expect(document.activeElement).toBe(root);

    root?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(root?.hidden).toBe(true);
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('detach removes the popover element from the host', () => {
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    expect(host.querySelector('.fc-quick')).not.toBeNull();
    handle.detach();
    expect(host.querySelector('.fc-quick')).toBeNull();
  });

  it('setStrings re-renders labels while open', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();
    const customTitle = 'CUSTOM_TITLE_X';
    handle.setStrings({
      ...defaultStrings,
      quickAnalysis: { ...defaultStrings.quickAnalysis, title: customTitle },
    });
    const titleEl = host.querySelector<HTMLElement>('.fc-quick__title');
    expect(titleEl?.textContent).toBe(customTitle);
    handle.detach();
  });
});
