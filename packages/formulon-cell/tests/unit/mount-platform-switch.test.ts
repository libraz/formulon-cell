import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
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

const LAYOUT_WRITES = [
  'setColumnWidth',
  'setRowHeight',
  'clearRowHeight',
  'clearColumnWidth',
  'setSheetFormatDefaults',
  'setWorkbookDefaultFont',
] as const;

/** Spy every layout/format writer the handle exposes; returns the spies. */
const spyLayoutWrites = (wb: WorkbookHandle): ReturnType<typeof vi.fn>[] => {
  const target = wb as unknown as Record<string, unknown>;
  const spies: ReturnType<typeof vi.fn>[] = [];
  for (const name of LAYOUT_WRITES) {
    if (typeof target[name] !== 'function') continue;
    spies.push(vi.spyOn(target as never, name as never) as unknown as ReturnType<typeof vi.fn>);
  }
  return spies;
};

describe('Spreadsheet platform switch', () => {
  beforeEach(() => {
    vi.stubGlobal('ResizeObserver', TestResizeObserver);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(makeCanvasContext());
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
    document.body.replaceChildren();
  });

  const mountMacJa = async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    expect(workbook.isStub).toBe(false);
    const instance = await Spreadsheet.mount(host, {
      workbook,
      locale: 'ja',
      ui: { platform: 'mac' },
    });
    return { workbook, instance };
  };

  it('re-projects the axis layout without writing layout to the engine', async () => {
    const { workbook, instance } = await mountMacJa();
    const layout = () => instance.store.getState().layout;
    expect(workbook.geometryPlatform).toBe('mac');
    expect(layout().defaultColWidth).toBe(100);
    expect(layout().screenScale).toBe(0.75);

    mutators.setColWidth(instance.store, 3, 120);
    instance.history.push({ undo: () => {}, redo: () => {} });
    expect(instance.history.canUndo()).toBe(true);
    const spies = spyLayoutWrites(workbook);
    expect(spies.length).toBeGreaterThan(0);

    instance.setUi({ platform: 'default' });
    expect(workbook.geometryPlatform).toBe('default');
    // Windows reads base 10 with 游ゴシック 12 (9 px MDW): 97 px rounded up to 104 px.
    expect(layout().defaultColWidth).toBe(104);
    expect(layout().screenScale).toBe(1);
    expect(instance.history.canUndo()).toBe(false);
    expect(instance.history.canRedo()).toBe(false);

    instance.setUi({ platform: 'mac' });
    expect(workbook.geometryPlatform).toBe('mac');
    expect(layout().defaultColWidth).toBe(100);
    expect(layout().screenScale).toBe(0.75);

    for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    instance.dispose();
  });

  it('leaves sizes alone on a theme-only setUi', async () => {
    const { instance } = await mountMacJa();
    mutators.setColWidth(instance.store, 2, 120);
    const before = instance.store.getState().layout;

    instance.setUi({ platform: 'mac', theme: 'ink' });
    const after = instance.store.getState().layout;
    expect(after.defaultColWidth).toBe(before.defaultColWidth);
    expect(after.defaultRowHeight).toBe(before.defaultRowHeight);
    expect(after.screenScale).toBe(before.screenScale);
    expect(after.colWidths.get(2)).toBe(before.colWidths.get(2));
    instance.dispose();
  });

  it('gives a mounted Mac workbook the platform even when the host created it for default', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const instance = await Spreadsheet.mount(host, { workbook, ui: { platform: 'mac' } });
    expect(workbook.geometryPlatform).toBe('mac');
    expect(instance.store.getState().layout.defaultColWidth).toBe(100);
    expect(instance.store.getState().layout.defaultRowHeight).toBeCloseTo(80 / 3, 6);
    instance.dispose();
  });
});
