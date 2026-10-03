import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import type { Extension, ExtensionHandle } from '../../src/extensions/types.js';
import { Spreadsheet } from '../../src/mount.js';
import { mutators } from '../../src/store/store.js';

class TestResizeObserver {
  static callbacks: ResizeObserverCallback[] = [];

  constructor(callback: ResizeObserverCallback) {
    TestResizeObserver.callbacks.push(callback);
  }

  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}

  static triggerLast(): void {
    const callback = TestResizeObserver.callbacks.at(-1);
    callback?.([], {} as ResizeObserver);
  }
}

const makeCanvasContext = (): CanvasRenderingContext2D =>
  new Proxy(
    {
      canvas: document.createElement('canvas'),
      measureText: (text: string) => ({ width: text.length * 7 }),
    },
    {
      get(target, prop) {
        if (prop in target) return target[prop as keyof typeof target];
        return vi.fn();
      },
      set(target, prop, value) {
        (target as Record<PropertyKey, unknown>)[prop] = value;
        return true;
      },
    },
  ) as unknown as CanvasRenderingContext2D;

describe('Spreadsheet feature registry', () => {
  beforeEach(() => {
    TestResizeObserver.callbacks = [];
    vi.stubGlobal('ResizeObserver', TestResizeObserver);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(makeCanvasContext());
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
    document.body.replaceChildren();
  });

  it('exposes built-in feature handles immediately after mount', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });

    const instance = await Spreadsheet.mount(host, { workbook });

    expect(instance.features.statusBar).toBeTruthy();
    // The flat view strip is opt-in; the ribbon's View tab covers it.
    expect(instance.features.viewToolbar).toBeUndefined();
    expect(instance.features.workbookObjects).toBeTruthy();
    expect(instance.features.clipboard).toBeTruthy();
    expect(instance.features.pasteSpecial).toBeTruthy();
    expect(instance.features.quickAnalysis).toBeTruthy();
    expect(instance.features.charts).toBeTruthy();
    expect(instance.features.pivotTableDialog).toBeTruthy();
    expect(instance.features.contextMenu).toBeTruthy();
    expect(instance.features.findReplace).toBeTruthy();
    expect(instance.features.validation).toBeTruthy();

    instance.dispose();
  });

  it('mounts user extensions after built-ins so ctx.resolve can compose them', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const resolved: Record<string, ExtensionHandle | undefined> = {};
    const ext: Extension = {
      id: 'probe',
      setup(ctx) {
        for (const id of ['clipboard', 'pasteSpecial', 'statusBar', 'workbookObjects']) {
          resolved[id] = ctx.resolve(id);
        }
        return { dispose() {} };
      },
    };

    const instance = await Spreadsheet.mount(host, { workbook, extensions: [ext] });

    expect(resolved.clipboard).toBeTruthy();
    expect(resolved.pasteSpecial).toBeTruthy();
    expect(resolved.statusBar).toBeTruthy();
    expect(resolved.workbookObjects).toBeTruthy();
    expect(instance.features.probe).toBeTruthy();

    instance.dispose();
  });

  it('delegates openPivotFieldList through a workbookObjects extension override', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const calls: string[] = [];
    const ext: Extension = {
      id: 'workbookObjects',
      setup() {
        return {
          dispose() {},
          openPivotFieldList(sheetIndex: number, pivotIndex: number) {
            calls.push(`${sheetIndex}:${pivotIndex}`);
            return true;
          },
        };
      },
    };

    const instance = await Spreadsheet.mount(host, { workbook, extensions: [ext] });

    expect(instance.openPivotFieldList(1, 2)).toBe(true);
    expect(calls).toEqual(['1:2']);
    instance.dispose();
  });

  it('returns false when opening a missing built-in PivotTable Field List', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, { workbook });

    expect(instance.openPivotFieldList(0, 0)).toBe(false);
    expect(instance.openActivePivotFieldList()).toBe(false);
    instance.dispose();
  });

  it('keeps an open PivotTable Field List synced to the active cell', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    vi.spyOn(workbook, 'getPivotTables').mockReturnValue([
      {
        sheetIndex: 0,
        pivotIndex: 0,
        top: 1,
        left: 1,
        rows: 2,
        cols: 2,
        cells: 4,
        fields: ['Region'],
        fieldItems: { Region: ['East'] },
      },
    ]);
    const instance = await Spreadsheet.mount(host, { workbook });

    expect(instance.openPivotFieldList(0, 0)).toBe(true);
    expect(host.querySelector('.fc-objects--taskpane')).toBeTruthy();

    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 2 });
    expect(host.querySelector('.fc-objects--taskpane')).toBeTruthy();

    mutators.setActive(instance.store, { sheet: 0, row: 9, col: 9 });
    expect(host.querySelector<HTMLElement>('.fc-objects')?.hidden).toBe(true);
    instance.dispose();
  });

  it('opens the built-in PivotTable Field List when selection enters a PivotTable', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    vi.spyOn(workbook, 'getPivotTables').mockReturnValue([
      {
        sheetIndex: 0,
        pivotIndex: 0,
        top: 2,
        left: 2,
        rows: 3,
        cols: 3,
        cells: 9,
        fields: ['Region'],
        fieldItems: { Region: ['East'] },
      },
    ]);
    const instance = await Spreadsheet.mount(host, { workbook });

    expect(host.querySelector('.fc-objects--taskpane')).toBeNull();

    mutators.setActive(instance.store, { sheet: 0, row: 3, col: 3 });

    expect(host.querySelector('.fc-objects--taskpane')).toBeTruthy();
    expect(host.querySelector('.fc-objects__title')?.textContent?.length).toBeGreaterThan(0);
    instance.dispose();
  });

  it('updates host-driven status bar upload and macro indicators through the instance API', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, {
      workbook,
      uploadStatus: 'saving',
      macroRecording: false,
    });
    instance.store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        statusOptions: { ...state.ui.statusOptions, uploadStatus: true, macroRecording: true },
      },
    }));

    instance.setUploadStatus('error');
    instance.setMacroRecording(true);

    expect(
      host.querySelector<HTMLElement>('.fc-host__statusbar-upload')?.dataset.uploadStatus,
    ).toBe('error');
    expect(
      host.querySelector<HTMLElement>('.fc-host__statusbar-macro')?.dataset.macroRecording,
    ).toBe('true');
    instance.dispose();
  });

  it('uses the next feature flags while attaching newly enabled host chrome', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, {
      workbook,
      features: { viewToolbar: false, workbookObjects: false },
    });

    instance.setFeatures({ viewToolbar: true, workbookObjects: true });

    expect(instance.features.viewToolbar).toBeTruthy();
    expect(instance.features.workbookObjects).toBeTruthy();
    expect(host.querySelector<HTMLButtonElement>('[aria-label="オブジェクト"]')).toBeTruthy();

    instance.dispose();
  });

  it('updates formula bar chrome labels when the locale changes', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, { workbook, locale: 'en' });

    expect(host.querySelector('[aria-label="Name box"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="Cancel formula edit"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="Enter formula"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="Formula bar"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="Expand formula bar"]')).toBeTruthy();
    const grid = host.querySelector<HTMLElement>('.fc-host__grid');
    const canvas = host.querySelector<HTMLCanvasElement>('.fc-host__canvas');
    const mirror = host.querySelector<HTMLElement>('.fc-host__a11y');
    const live = host.querySelector<HTMLElement>('.fc-host__a11y-live');
    expect(grid?.getAttribute('role')).toBe('grid');
    expect(grid?.getAttribute('aria-label')).toBe('Worksheet grid');
    expect(grid?.tabIndex).toBe(-1);
    expect(grid?.getAttribute('aria-describedby')).toBe(live?.id);
    expect(grid?.getAttribute('aria-activedescendant')).toBe(`${mirror?.id}-active-cell`);
    expect(grid?.getAttribute('aria-rowcount')).toBe('1048576');
    expect(grid?.getAttribute('aria-colcount')).toBe('16384');
    expect(canvas?.getAttribute('aria-hidden')).toBe('true');
    // The mirror sits inside the grid as grid > rowgroup > row > gridcell; the
    // announcer is a separate node outside the grid.
    expect(mirror?.getAttribute('role')).toBe('rowgroup');
    expect(mirror?.parentElement).toBe(grid);
    expect(mirror?.getAttribute('aria-live')).toBeNull();
    expect(live?.getAttribute('aria-live')).toBe('polite');
    expect(live?.getAttribute('aria-atomic')).toBe('true');
    expect(live?.getAttribute('role')).toBeNull();
    expect(live?.closest('.fc-host__grid')).toBeNull();
    const activeCell = mirror?.querySelector<HTMLElement>('[aria-selected="true"]');
    expect(activeCell?.id).toBe(`${mirror?.id}-active-cell`);
    expect(activeCell?.getAttribute('role')).toBe('gridcell');
    expect(activeCell?.parentElement?.getAttribute('role')).toBe('row');
    expect(activeCell?.parentElement?.getAttribute('aria-rowindex')).toBe('1');

    instance.i18n.setLocale('ja');

    expect(host.querySelector('[aria-label="名前ボックス"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式の編集をキャンセル"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式を入力"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式バー"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式バーを展開"]')).toBeTruthy();
    expect(grid?.getAttribute('aria-label')).toBe('ワークシート グリッド');

    instance.dispose();
  });

  it('keeps the Mac draft mirror anchored through format, freeze, scroll, RTL, and resize changes', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, {
      workbook,
      ui: { platform: 'mac' },
      features: { fxDialog: true },
    });
    const grid = host.querySelector<HTMLElement>('.fc-host__grid');
    const mirror = host.querySelector<HTMLElement>('.fc-host__formula-draft-mirror');
    if (!grid || !mirror) throw new Error('expected Mac grid mirror');
    Object.defineProperty(grid, 'clientWidth', { configurable: true, value: 500 });
    Object.defineProperty(grid, 'clientHeight', { configurable: true, value: 400 });
    vi.spyOn(grid, 'getBoundingClientRect').mockReturnValue({
      x: 0,
      y: 0,
      top: 0,
      left: 0,
      right: 500,
      bottom: 400,
      width: 500,
      height: 400,
      toJSON: () => ({}),
    } as DOMRect);
    mutators.setViewportSize(instance.store, 20, 12, 500);
    const anchor = { sheet: 0, row: 2, col: 2 };
    mutators.setActive(instance.store, anchor);
    mutators.setFreezePanes(instance.store, 1, 1);
    mutators.setCellFormat(instance.store, anchor, {
      fill: '#123456',
      color: '#abcdef',
      bold: true,
    });
    instance.openFunctionArguments('SUM');

    const state = instance.store.getState();
    expect(mirror.getAttribute('aria-hidden')).toBe('true');
    expect(mirror.hidden).toBe(false);
    expect(mirror.textContent?.startsWith('=')).toBe(true);
    expect(mirror.style.left).toBe(
      `${state.layout.headerColWidth + state.layout.defaultColWidth * 2}px`,
    );
    expect(mirror.style.top).toBe(
      `${state.layout.headerRowHeight + state.layout.defaultRowHeight * 2}px`,
    );
    expect(mirror.style.background).not.toBe('');
    expect(mirror.style.color).not.toBe('');

    mutators.scrollBy(instance.store, 2, 2);
    expect(mirror.hidden).toBe(true);
    mutators.scrollBy(instance.store, -2, -2);
    expect(mirror.hidden).toBe(false);
    mutators.setRightToLeft(instance.store, true);
    expect(mirror.style.direction).toBe('rtl');
    expect(mirror.hidden).toBe(false);
    TestResizeObserver.triggerLast();
    expect(mirror.hidden).toBe(false);
    expect(workbook.cellFormula(anchor)).toBeNull();
    expect(instance.history.canUndo()).toBe(false);

    instance.dispose();
  });

  it('routes Mac picker clicks to selection, argument clicks to the palette target, and keeps inline editing first', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, {
      workbook,
      ui: { platform: 'mac' },
      features: { fxDialog: true },
    });
    const grid = host.querySelector<HTMLElement>('.fc-host__grid');
    if (!grid) throw new Error('expected grid');
    mutators.setViewportSize(instance.store, 20, 12, 500);
    const cell = (row: number, col: number): { clientX: number; clientY: number } => ({
      clientX: 26 + 75 * col + 10,
      clientY: 20 + 20 * row + 10,
    });
    const fireDown = (row: number, col: number): PointerEvent => {
      const event = new PointerEvent('pointerdown', {
        ...cell(row, col),
        button: 0,
        bubbles: true,
        cancelable: true,
        pointerId: 1,
      });
      grid.dispatchEvent(event);
      return event;
    };

    instance.openFunctionArguments();
    fireDown(1, 1);
    expect(instance.store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 1 });

    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });
    instance.openFunctionArguments('SUM');
    const argumentPointer = fireDown(1, 1);
    expect(argumentPointer.defaultPrevented).toBe(true);
    expect(instance.store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input')?.value).toContain(
      'B2',
    );

    workbook.setFormula({ sheet: 0, row: 0, col: 0 }, '=A2');
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });
    grid.dispatchEvent(
      new MouseEvent('dblclick', {
        ...cell(0, 0),
        button: 0,
        bubbles: true,
        cancelable: true,
      }),
    );
    const inlineEditor = host.querySelector<HTMLTextAreaElement>('.fc-host__editor');
    expect(inlineEditor).not.toBeNull();
    const beforeInlineValue = inlineEditor?.value;
    instance.openFunctionArguments();
    const inlinePointer = fireDown(1, 1);
    expect(inlinePointer.defaultPrevented).toBe(true);
    expect(inlineEditor?.value).not.toBe(beforeInlineValue);
    expect(inlineEditor?.value).toContain('B2');
    expect(instance.store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });

    instance.dispose();
  });

  it('cancels the Mac pane across sheet, platform, workbook, and dispose transitions', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const oldWorkbook = await WorkbookHandle.createDefault();
    expect(oldWorkbook.isStub).toBe(false);
    const anchor = { sheet: 0, row: 0, col: 0 };
    oldWorkbook.setNumber(anchor, 9);
    const instance = await Spreadsheet.mount(host, {
      workbook: oldWorkbook,
      ui: { platform: 'mac' },
    });
    mutators.replaceCells(instance.store, oldWorkbook.cells(0));
    mutators.setActive(instance.store, anchor);
    instance.history.clear();
    instance.openFunctionArguments('SUM');
    const macRoot = host.querySelector<HTMLElement>('.fc-mac-formula-palette');
    const dock = host.querySelector<HTMLElement>('.fc-host__taskpane-dock');
    expect(macRoot?.hidden).toBe(false);
    expect(oldWorkbook.cellFormula(anchor)).toBeNull();
    expect(instance.history.canUndo()).toBe(false);

    mutators.setSheetIndex(instance.store, 1);
    expect(macRoot?.hidden).toBe(true);
    expect(oldWorkbook.getValue(anchor)).toEqual({ kind: 'number', value: 9 });
    expect(instance.history.canUndo()).toBe(false);

    instance.setUi({ platform: 'default' });
    expect(host.querySelector('.fc-mac-formula-palette')).toBeNull();
    expect(document.querySelector('.fc-fxdialog')).toBeTruthy();
    instance.setUi({ platform: 'mac' });
    expect(host.querySelector('.fc-mac-formula-palette')).toBeTruthy();
    mutators.setSheetIndex(instance.store, 0);
    mutators.replaceCells(instance.store, oldWorkbook.cells(0));
    mutators.setActive(instance.store, anchor);
    instance.openFunctionArguments('SUM');

    const nextWorkbook = await WorkbookHandle.createDefault();
    expect(nextWorkbook.isStub).toBe(false);
    await instance.setWorkbook(nextWorkbook);
    expect(host.querySelector<HTMLElement>('.fc-mac-formula-palette')?.hidden).toBe(true);
    expect(oldWorkbook.getValue(anchor)).toEqual({ kind: 'number', value: 9 });
    expect(nextWorkbook.cellFormula(anchor)).toBeNull();
    expect(instance.history.canUndo()).toBe(false);

    instance.openFunctionArguments('SUM');
    expect(nextWorkbook.cellFormula(anchor)).toBeNull();
    expect(instance.history.canUndo()).toBe(false);
    instance.dispose();
    expect(dock?.hidden).toBe(true);
  });

  it('setTheme updates host theme state, store state, and emits themeChange', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, { workbook, theme: 'paper' });
    const onThemeChange = vi.fn();
    const unsubscribe = instance.on('themeChange', onThemeChange);

    expect(host.dataset.fcTheme).toBe('paper');
    expect(instance.store.getState().ui.theme).toBe('paper');

    instance.setTheme('ink');

    expect(host.dataset.fcTheme).toBe('ink');
    expect(instance.store.getState().ui.theme).toBe('ink');
    expect(onThemeChange).toHaveBeenCalledTimes(1);
    expect(onThemeChange).toHaveBeenCalledWith({ theme: 'ink' });

    unsubscribe();
    instance.setTheme('contrast');
    expect(onThemeChange).toHaveBeenCalledTimes(1);

    instance.dispose();
  });
});
