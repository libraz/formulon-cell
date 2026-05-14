import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import type { Extension, ExtensionHandle } from '../../src/extensions/types.js';
import { Spreadsheet } from '../../src/mount.js';

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

    instance.i18n.setLocale('ja');

    expect(host.querySelector('[aria-label="名前ボックス"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式の編集をキャンセル"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式を入力"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式バー"]')).toBeTruthy();
    expect(host.querySelector('[aria-label="数式バーを展開"]')).toBeTruthy();

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
