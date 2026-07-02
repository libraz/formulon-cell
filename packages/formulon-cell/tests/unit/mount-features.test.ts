import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import type { Extension, ExtensionHandle } from '../../src/extensions/types.js';
import { Spreadsheet } from '../../src/mount.js';
import { mutators } from '../../src/store/store.js';

class TestResizeObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
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
    expect(instance.features.viewToolbar).toBeTruthy();
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
    const live = host.querySelector<HTMLElement>('.fc-host__a11y');
    expect(grid?.getAttribute('role')).toBe('grid');
    expect(grid?.getAttribute('aria-label')).toBe('Worksheet grid');
    expect(grid?.tabIndex).toBe(-1);
    expect(grid?.getAttribute('aria-describedby')).toBe(live?.id);
    expect(grid?.getAttribute('aria-activedescendant')).toBe(`${live?.id}-active-cell`);
    expect(grid?.getAttribute('aria-rowcount')).toBe('1048576');
    expect(grid?.getAttribute('aria-colcount')).toBe('16384');
    expect(canvas?.getAttribute('aria-hidden')).toBe('true');
    expect(live?.getAttribute('aria-live')).toBe('polite');
    expect(live?.getAttribute('aria-atomic')).toBe('true');
    const activeCell = live?.querySelector<HTMLElement>('[role="gridcell"]');
    expect(activeCell?.id).toBe(`${live?.id}-active-cell`);
    expect(activeCell?.getAttribute('aria-selected')).toBe('true');

    instance.i18n.setLocale('ja');

    expect(host.querySelector('[aria-label="名前ボックス"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式の編集をキャンセル"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式を入力"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式バー"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式バーを展開"]')).toBeTruthy();
    expect(grid?.getAttribute('aria-label')).toBe('ワークシート グリッド');

    instance.dispose();
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
