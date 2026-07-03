import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachQuickAnalysis } from '../../../src/interact/quick-analysis.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

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

  it('open() reveals the popover and renders Excel-style group tabs', () => {
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
    const tabs = root?.querySelectorAll('[role="tab"]') ?? [];
    expect(tabs.length).toBeGreaterThanOrEqual(3);
    expect(root?.querySelector('[role="tab"][aria-selected="true"]')?.textContent).toBe(
      defaultStrings.quickAnalysis.groups.formatting,
    );
    handle.close();
    expect(root?.hidden).toBe(true);
    handle.detach();
  });

  it('shows an Excel-style Quick Analysis button for a multi-cell selection', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
    ]);
    setRange(store, 0, 0, 0, 1);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    const button = host.querySelector<HTMLButtonElement>('.fc-quick__button');

    expect(button).not.toBeNull();
    expect(button?.hidden).toBe(false);
    expect(button?.getAttribute('aria-label')).toBe(defaultStrings.quickAnalysis.title);

    button?.click();
    expect(host.querySelector<HTMLElement>('.fc-quick')?.hidden).toBe(false);
    expect(button?.hidden).toBe(true);
    handle.detach();
  });

  it('hides the Quick Analysis button for a single-cell selection', () => {
    seed(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });

    expect(host.querySelector<HTMLButtonElement>('.fc-quick__button')?.hidden).toBe(true);
    handle.detach();
  });

  it('hides the Quick Analysis button while editing and restores it when editing ends', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
    ]);
    setRange(store, 0, 0, 0, 1);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    const button = host.querySelector<HTMLButtonElement>('.fc-quick__button');
    expect(button?.hidden).toBe(false);

    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, editor: { kind: 'edit', raw: '1', caret: 1 } },
    }));
    expect(button?.hidden).toBe(true);

    store.setState((s) => ({ ...s, ui: { ...s.ui, editor: { kind: 'idle' } } }));
    expect(button?.hidden).toBe(false);
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

  it('projects disabled reasons on unavailable Quick Analysis actions', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 'A' },
      { row: 0, col: 1, value: 'B' },
    ]);
    setRange(store, 0, 0, 0, 1);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();

    const dataBar = host.querySelector<HTMLButtonElement>('button[data-action="format-data-bar"]');
    const pivot = host.querySelector<HTMLButtonElement>('button[data-action="tables-pivot"]');
    const chart = host.querySelector<HTMLButtonElement>('button[data-action="charts-column"]');

    expect(dataBar?.disabled).toBe(true);
    expect(dataBar?.dataset.disabledReason).toBe(
      defaultStrings.quickAnalysis.disabledReasons.requiresTwoNumbers,
    );
    expect(pivot?.disabled).toBe(true);
    expect(pivot?.dataset.disabledReason).toBe(
      defaultStrings.quickAnalysis.disabledReasons.pivotUnavailable,
    );
    expect(chart?.disabled).toBe(true);
    expect(chart?.getAttribute('aria-description')).toBe(
      defaultStrings.quickAnalysis.disabledReasons.requiresNumbers,
    );
    handle.detach();
  });

  it('switches visible Quick Analysis tab panels when a group tab is clicked', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 0, col: 2, value: 3 },
    ]);
    setRange(store, 0, 0, 0, 2);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();

    const totalsTab = host.querySelector<HTMLButtonElement>('[role="tab"][data-group="totals"]');
    totalsTab?.click();
    const selectedTotalsTab = host.querySelector<HTMLButtonElement>(
      '[role="tab"][data-group="totals"]',
    );

    expect(selectedTotalsTab?.getAttribute('aria-selected')).toBe('true');
    expect(host.querySelector<HTMLElement>('#fc-quick-panel-formatting')?.hidden).toBe(true);
    expect(host.querySelector<HTMLElement>('#fc-quick-panel-totals')?.hidden).toBe(false);
    handle.detach();
  });

  it('Quick Analysis tabs support Excel-style arrow, Home, and End navigation', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 0, col: 2, value: 3 },
    ]);
    setRange(store, 0, 0, 0, 2);
    const handle = attachQuickAnalysis({ host, store, wb, strings: defaultStrings });
    handle.open();

    const selected = () =>
      host.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]');
    selected()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(selected()?.dataset.group).toBe('charts');

    selected()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(selected()?.dataset.group).toBe('sparklines');

    selected()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(selected()?.dataset.group).toBe('formatting');
    expect(document.activeElement).toBe(selected());
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

  it('opens the PivotTable creation flow from the Tables group when available', () => {
    seed(store, wb, [
      { row: 0, col: 0, value: 'Region' },
      { row: 0, col: 1, value: 'Sales' },
      { row: 1, col: 0, value: 'East' },
      { row: 1, col: 1, value: 12 },
    ]);
    Object.defineProperty(wb, 'capabilities', {
      configurable: true,
      value: { ...wb.capabilities, pivotTableMutate: true },
    });
    setRange(store, 0, 0, 1, 1);
    let opened = 0;
    const handle = attachQuickAnalysis({
      host,
      store,
      wb,
      strings: defaultStrings,
      onOpenPivotTable: () => {
        opened += 1;
      },
      canOpenPivotTable: () => true,
    });
    handle.open();

    host.querySelector<HTMLButtonElement>('[role="tab"][data-group="tables"]')?.click();
    const pivotBtn = host.querySelector<HTMLButtonElement>('button[data-action="tables-pivot"]');
    expect(pivotBtn).not.toBeNull();
    expect(pivotBtn?.disabled).toBe(false);
    expect(pivotBtn?.textContent).toBe(defaultStrings.quickAnalysis.actions.pivotTable);

    pivotBtn?.click();

    expect(opened).toBe(1);
    expect(host.querySelector<HTMLElement>('.fc-quick')?.hidden).toBe(true);
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
    expect(host.querySelector('.fc-quick__button')).toBeNull();
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

  it('keeps launcher, tab, and action button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/quick-analysis.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('function createQuickAnalysisLauncher');
    expect(source).toContain('function createQuickAnalysisTab');
    expect(source).toContain('function createQuickAnalysisActionButton');
    expect(source).toContain('const button = createQuickAnalysisLauncher()');
    expect(source).toContain('const tab = createQuickAnalysisTab(');
    expect(source).toContain('const btn = createQuickAnalysisActionButton(strings, action)');
    expect(source).toContain('projectDisabledState(button, disabled');
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps Quick Analysis close to Excel 365 desktop gallery geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/overlays/quick-analysis-and-charts.css'),
      'utf8',
    );
    const quickCss = css.slice(css.indexOf('.fc-quick {'), css.indexOf('.fc-charts {'));

    expect(css).toMatch(
      /\.fc-quick__button\s*\{[\s\S]*?width: 22px;[\s\S]*?height: 22px;[\s\S]*?border-radius: 2px;[\s\S]*?0 3px 8px rgba\(0, 0, 0, 0\.16\)/,
    );
    expect(quickCss).toMatch(
      /\.fc-quick\s*\{[\s\S]*?padding: 6px;[\s\S]*?border-radius: 2px;[\s\S]*?box-shadow:/,
    );
    expect(quickCss).toMatch(
      /\.fc-quick__tab\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 4px 8px 5px;/,
    );
    expect(quickCss).toMatch(
      /\.fc-quick__tab:hover,[\s\S]*?\.fc-quick__tab:focus-visible\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(quickCss).toMatch(
      /\.fc-quick__action\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 7px;[\s\S]*?border-radius: 2px;/,
    );
    expect(quickCss).toMatch(
      /\.fc-quick__action:hover:not\(:disabled\),[\s\S]*?\.fc-quick__action:focus-visible:not\(:disabled\)\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(quickCss).not.toContain('border-radius: var(--fc-radius-md, 6px);');
    expect(quickCss).not.toContain('background: var(--fc-accent-soft');
    expect(css).not.toContain('box-shadow: var(--fc-shadow-2)');
  });
});
