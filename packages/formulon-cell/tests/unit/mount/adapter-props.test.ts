import { describe, expect, it, vi } from 'vitest';
import {
  applyControlledLocale,
  applyControlledProps,
  buildAdapterRibbonWiring,
} from '../../../src/mount/adapter-props.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';

const fakeInstance = (over: Record<string, unknown> = {}): SpreadsheetInstance =>
  ({
    locale: 'ja',
    workbook: {},
    i18n: { locale: 'ja', strings: {}, extend: vi.fn() },
    setLocale: vi.fn(() => true),
    setTheme: vi.fn(),
    setUi: vi.fn(),
    setFeatures: vi.fn(),
    setToolbar: vi.fn(),
    setPolicy: vi.fn(),
    setViewportOptions: vi.fn(),
    setContextMenu: vi.fn(),
    setOverlayOptions: vi.fn(),
    setExtensions: vi.fn(),
    setPrinterProfiles: vi.fn(),
    setPrinterProfileId: vi.fn(),
    setUploadStatus: vi.fn(),
    setMacroRecording: vi.fn(),
    setWorkbook: vi.fn(async () => {}),
    ...over,
  }) as unknown as SpreadsheetInstance;

describe('adapter-props applyControlledLocale', () => {
  it('reports a rejected locale through both handlers', () => {
    const inst = fakeInstance({ setLocale: vi.fn(() => false) });
    const onLocaleChange = vi.fn();
    const onError = vi.fn();
    expect(applyControlledLocale(inst, 'xx', { onLocaleChange, onError })).toBe(false);
    expect(onLocaleChange).toHaveBeenCalledWith({ locale: 'ja', strings: {} });
    expect(onError).toHaveBeenCalledOnce();
  });

  it('skips the fallback event when the core already emitted one', () => {
    let gen = 0;
    const inst = fakeInstance({
      setLocale: vi.fn(() => {
        gen += 1;
        return false;
      }),
    });
    const onLocaleChange = vi.fn();
    applyControlledLocale(inst, 'xx', { onLocaleChange, getLocaleEventGeneration: () => gen });
    expect(onLocaleChange).not.toHaveBeenCalled();
  });
});

describe('adapter-props applyControlledProps', () => {
  it('applies only props that differ from the baseline', async () => {
    const inst = fakeInstance();
    const policy = {};
    await applyControlledProps(inst, { policy, macroRecording: true }, { policy });
    expect(inst.setPolicy).not.toHaveBeenCalled();
    expect(inst.setMacroRecording).toHaveBeenCalledWith(true);
  });
});

describe('adapter-props buildAdapterRibbonWiring', () => {
  it('registers overrides and hooks only for provided callbacks', () => {
    const wiring = buildAdapterRibbonWiring(fakeInstance(), () => ({ onRunScript: vi.fn() }));
    expect(wiring.dynamicDropdowns.applyScriptAction).toBeTypeOf('function');
    expect(wiring.dynamicDropdowns.applyAddInAction).toBeUndefined();
    expect(wiring.hooks.review).toEqual({});
  });

  it('routes custom script action to the latest callback and ink modes to draw hooks', () => {
    let cbs: { onRunScript?: () => void; onDrawPen?: () => void; onDrawEraser?: () => void } = {
      onRunScript: vi.fn(),
    };
    const first = cbs.onRunScript;
    const wiring = buildAdapterRibbonWiring(fakeInstance(), () => cbs);
    cbs = { onRunScript: vi.fn(), onDrawPen: vi.fn(), onDrawEraser: vi.fn() };
    wiring.dynamicDropdowns.applyScriptAction?.('custom');
    expect(first).not.toHaveBeenCalled();
    expect(cbs.onRunScript).toHaveBeenCalledOnce();
    wiring.hooks.drawing?.setInkMode?.('pen');
    wiring.hooks.drawing?.setInkMode?.('erase');
    expect(cbs.onDrawPen).toHaveBeenCalledOnce();
    expect(cbs.onDrawEraser).toHaveBeenCalledOnce();
  });

  it('lets host dropdownActions win', () => {
    const applyScriptAction = vi.fn();
    const wiring = buildAdapterRibbonWiring(fakeInstance(), () => ({ onRunScript: vi.fn() }), {
      applyScriptAction,
    });
    expect(wiring.dynamicDropdowns.applyScriptAction).toBe(applyScriptAction);
  });
});
