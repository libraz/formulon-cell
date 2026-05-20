import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachViewToolbar } from '../../../src/interact/view-toolbar.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const fakeWb = () =>
  ({
    capabilities: {
      colRowSize: false,
      hiddenRowsCols: false,
      outlines: false,
      freeze: true,
      sheetTabHidden: false,
    },
    setSheetFreeze: vi.fn(),
    setSheetZoom: vi.fn(),
  }) as unknown as WorkbookHandle & {
    setSheetFreeze: ReturnType<typeof vi.fn>;
    setSheetZoom: ReturnType<typeof vi.fn>;
  };

describe('attachViewToolbar', () => {
  let toolbar: HTMLElement;
  let store: ReturnType<typeof createSpreadsheetStore>;
  let history: History;

  beforeEach(() => {
    toolbar = document.createElement('div');
    document.body.appendChild(toolbar);
    store = createSpreadsheetStore();
    history = new History();
  });

  it('renders spreadsheet-style view toggles and applies them to the store', () => {
    const wb = fakeWb();
    const invalidate = vi.fn();
    const handle = attachViewToolbar({
      toolbar,
      store,
      wb,
      history,
      strings: en,
      onChanged: invalidate,
    });

    const gridlines = toolbar.querySelector<HTMLButtonElement>('button[aria-label="Gridlines"]');
    const formulas = toolbar.querySelector<HTMLButtonElement>('button[aria-label="Formulas"]');
    const r1c1 = toolbar.querySelector<HTMLButtonElement>('button[aria-label="R1C1"]');
    const pageBreak = toolbar.querySelector<HTMLButtonElement>(
      'button[aria-label="Page Break Preview"]',
    );
    const normal = toolbar.querySelector<HTMLButtonElement>('button[aria-label="Normal"]');
    expect(normal?.getAttribute('aria-pressed')).toBe('true');
    expect(gridlines?.getAttribute('aria-pressed')).toBe('true');

    pageBreak?.click();
    gridlines?.click();
    formulas?.click();
    r1c1?.click();

    expect(store.getState().ui.workbookView).toBe('pageBreakPreview');
    expect(pageBreak?.getAttribute('aria-pressed')).toBe('true');
    expect(normal?.getAttribute('aria-pressed')).toBe('false');
    expect(store.getState().ui.showGridLines).toBe(false);
    expect(store.getState().ui.showFormulas).toBe(true);
    expect(store.getState().ui.r1c1).toBe(true);
    expect(invalidate).toHaveBeenCalledTimes(4);
    handle.detach();
  });

  it('drives freeze panes and persists the workbook view when supported', () => {
    const wb = fakeWb();
    mutators.setActive(store, { sheet: 0, row: 3, col: 2 });
    const handle = attachViewToolbar({ toolbar, store, wb, history, strings: en });

    toolbar.querySelector<HTMLButtonElement>('button[aria-label="Freeze Panes"]')?.click();
    expect(store.getState().layout.freezeRows).toBe(3);
    expect(store.getState().layout.freezeCols).toBe(2);
    expect(wb.setSheetFreeze).toHaveBeenLastCalledWith(0, 3, 2);

    toolbar.querySelector<HTMLButtonElement>('button[aria-label="Unfreeze"]')?.click();
    expect(store.getState().layout.freezeRows).toBe(0);
    expect(store.getState().layout.freezeCols).toBe(0);
    expect(wb.setSheetFreeze).toHaveBeenLastCalledWith(0, 0, 0);
    handle.detach();
  });

  it('sets zoom through the workbook and relabels on locale changes', () => {
    const wb = fakeWb();
    const openObjects = vi.fn();
    const handle = attachViewToolbar({
      toolbar,
      store,
      wb,
      history,
      strings: en,
      onOpenObjects: openObjects,
    });
    const select = toolbar.querySelector<HTMLSelectElement>('.fc-viewbar__select');
    if (!select) throw new Error('expected zoom select');
    expect(Array.from(select.options).map((option) => option.value)).toEqual([
      '75',
      '100',
      '125',
      '150',
      '200',
    ]);

    select.value = '150';
    select.dispatchEvent(new Event('change'));
    expect(store.getState().viewport.zoom).toBe(1.5);
    expect(wb.setSheetZoom).toHaveBeenLastCalledWith(0, 150);

    handle.setStrings(ja);
    expect(toolbar.textContent).toContain('表示');
    expect(
      toolbar.querySelector<HTMLButtonElement>('button[aria-label="改ページ プレビュー"]'),
    ).not.toBeNull();
    expect(toolbar.querySelector<HTMLButtonElement>('button[aria-label="枠線"]')).not.toBeNull();
    toolbar.querySelector<HTMLButtonElement>('button[aria-label="オブジェクト"]')?.click();
    expect(openObjects).toHaveBeenCalledOnce();
    handle.detach();
  });

  it('omits the workbook objects button when no opener is supplied', () => {
    const wb = fakeWb();
    const handle = attachViewToolbar({ toolbar, store, wb, history, strings: en });

    expect(toolbar.querySelector<HTMLButtonElement>('button[aria-label="Objects"]')).toBeNull();
    handle.detach();
  });

  it('saves, activates, and deletes sheet views from the toolbar', () => {
    const wb = fakeWb();
    const invalidate = vi.fn();
    const handle = attachViewToolbar({
      toolbar,
      store,
      wb,
      history,
      strings: en,
      onChanged: invalidate,
    });

    mutators.setFilterRange(store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
    mutators.setFreezePanes(store, 2, 1);
    const deleteButton = toolbar.querySelector<HTMLButtonElement>('button[aria-label="Delete"]');
    expect(deleteButton?.disabled).toBe(true);
    expect(deleteButton?.dataset.disabledReason).toBe(en.viewToolbar.deleteViewRequiresActive);
    expect(deleteButton?.getAttribute('aria-description')).toBe(
      en.viewToolbar.deleteViewRequiresActive,
    );
    expect(deleteButton?.title).toBe(
      `${en.viewToolbar.deleteView}\n${en.viewToolbar.deleteViewRequiresActive}`,
    );
    toolbar.querySelector<HTMLButtonElement>('button[aria-label="Save"]')?.click();
    const saved = store.getState().sheetViews.views[0];
    expect(saved).toMatchObject({ name: 'Views 1', freeze: { rows: 2, cols: 1 } });

    mutators.setFilterRange(store, null);
    mutators.setFreezePanes(store, 0, 0);
    const select = toolbar.querySelector<HTMLSelectElement>('select[aria-label="Views"]');
    if (!select || !saved) throw new Error('missing sheet-view controls');
    expect(Array.from(select.options).map((option) => option.textContent)).toEqual([
      'Current view',
      'Views 1',
    ]);
    select.value = saved.id;
    select.dispatchEvent(new Event('change'));
    expect(store.getState().sheetViews.activeViewId).toBe(saved.id);
    expect(store.getState().layout.freezeRows).toBe(2);
    expect(store.getState().ui.filterRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
    expect(deleteButton?.disabled).toBe(false);
    expect(deleteButton?.dataset.disabledReason).toBeUndefined();
    expect(deleteButton?.hasAttribute('aria-description')).toBe(false);
    expect(deleteButton?.title).toBe(en.viewToolbar.deleteView);

    deleteButton?.click();
    expect(store.getState().sheetViews.views).toEqual([]);
    expect(store.getState().sheetViews.activeViewId).toBeNull();
    expect(deleteButton?.disabled).toBe(true);
    expect(deleteButton?.dataset.disabledReason).toBe(en.viewToolbar.deleteViewRequiresActive);
    expect(invalidate).toHaveBeenCalledTimes(3);
    handle.detach();
  });

  it('keeps view toolbar button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/view-toolbar.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('function createViewToolbarButton');
    expect(source).toContain("const normalView = createViewToolbarButton('fc-viewbar__button')");
    expect(source).toContain("const gridlines = createViewToolbarButton('fc-viewbar__toggle')");
    expect(source).toContain("const zoomFit = createViewToolbarButton('fc-viewbar__button')");
    expect(source).toContain("const objectsBtn = createViewToolbarButton('fc-viewbar__button')");
    expect(source).not.toContain('function makeButton');
    expect(source).not.toContain("document.createElement('button')");
  });
});
