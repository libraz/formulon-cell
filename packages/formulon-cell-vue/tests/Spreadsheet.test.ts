import * as Core from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { nextTick } from 'vue';
import type { ScreenClipCapture, ScreenClipResult } from '../src';
import * as VuePackage from '../src';
import { type MountedVueSpreadsheet, mountVueSpreadsheet } from './test-utils/mount';

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
  await nextTick();
};

/**
 * End-to-end tests for the Vue `<Spreadsheet>` wrapper. Like the React side,
 * we deliberately avoid mocking `Spreadsheet.mount` and exercise the real
 * core (with the stub WASM engine) so the prop / event plumbing is verified
 * against the live instance.
 */
describe('Vue <Spreadsheet>', () => {
  let mounted: MountedVueSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) {
      await mounted.dispose();
      mounted = null;
    }
    document.body.replaceChildren();
  });

  it('emits ready with a live SpreadsheetInstance after mount completes', async () => {
    const onReady = vi.fn();
    mounted = await mountVueSpreadsheet({ listeners: { onReady } });

    expect(onReady).toHaveBeenCalledTimes(1);
    const inst = onReady.mock.calls[0]?.[0];
    expect(inst).toBeDefined();
    expect(typeof inst?.dispose).toBe('function');
    expect(typeof inst?.setTheme).toBe('function');
    // The exposed ref should resolve to the same instance.
    expect(mounted.exposed.instance.value).toBe(inst);
  });

  it('emits cellChange when the underlying engine reports a cell mutation', async () => {
    const onCellChange = vi.fn();
    mounted = await mountVueSpreadsheet({ listeners: { onCellChange } });

    mounted.instance.workbook.setNumber({ sheet: 0, row: 2, col: 3 }, 11);
    await flush();

    expect(onCellChange).toHaveBeenCalledTimes(1);
    const event = onCellChange.mock.calls[0]?.[0];
    expect(event?.addr).toEqual({ sheet: 0, row: 2, col: 3 });
    expect(event?.value).toEqual({ kind: 'number', value: 11 });
  });

  it('emits selectionChange when the active selection moves', async () => {
    const onSelectionChange = vi.fn();
    mounted = await mountVueSpreadsheet({ listeners: { onSelectionChange } });

    Core.mutators.setActive(mounted.instance.store, { sheet: 0, row: 4, col: 1 });
    await flush();

    expect(onSelectionChange).toHaveBeenCalled();
    const last = onSelectionChange.mock.calls.at(-1)?.[0];
    expect(last?.active).toEqual({ sheet: 0, row: 4, col: 1 });
    expect(last?.range).toMatchObject({ sheet: 0, r0: 4, c0: 1, r1: 4, c1: 1 });
  });

  it('emits workbookChange when the workbook prop swaps to a new handle', async () => {
    const onWorkbookChange = vi.fn();
    mounted = await mountVueSpreadsheet({ listeners: { onWorkbookChange } });
    const original = mounted.instance.workbook;

    const next = await Core.WorkbookHandle.createDefault({ preferStub: true });
    expect(next).not.toBe(original);

    await mounted.setProp('workbook', next);
    // setWorkbook is async on the instance side — drain a few ticks.
    for (let i = 0; i < 10; i += 1) {
      if (onWorkbookChange.mock.calls.length > 0) break;
      await flush();
    }

    expect(onWorkbookChange).toHaveBeenCalled();
    const event = onWorkbookChange.mock.calls.at(-1)?.[0];
    expect(event?.workbook).toBe(next);
    expect(mounted.instance.workbook).toBe(next);
  });

  it('forwards a theme prop change to instance.setTheme without remounting', async () => {
    const onThemeChange = vi.fn();
    mounted = await mountVueSpreadsheet({
      props: { theme: 'paper' },
      listeners: { onThemeChange },
    });
    const original = mounted.instance;
    expect(mounted.host.querySelector<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme).toBe(
      'paper',
    );

    await mounted.setProp('theme', 'ink');

    expect(mounted.exposed.instance.value).toBe(original);
    expect(onThemeChange).toHaveBeenCalledTimes(1);
    expect(onThemeChange.mock.calls[0]?.[0]).toEqual({ theme: 'ink' });
    expect(mounted.host.querySelector<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme).toBe('ink');
  });

  it('forwards a locale prop change to instance.i18n.setLocale without remounting', async () => {
    const onLocaleChange = vi.fn();
    mounted = await mountVueSpreadsheet({
      props: { locale: 'en' },
      listeners: { onLocaleChange },
    });
    const original = mounted.instance;
    expect(mounted.instance.i18n.locale).toBe('en');

    await mounted.setProp('locale', 'ja');

    expect(mounted.exposed.instance.value).toBe(original);
    expect(mounted.instance.i18n.locale).toBe('ja');
    expect(onLocaleChange).toHaveBeenCalled();
    const event = onLocaleChange.mock.calls.at(-1)?.[0];
    expect(event?.locale).toBe('ja');
  });

  it('keeps the Screen Clipping capture hook current without remounting', async () => {
    const firstCapture = vi.fn(() => 'data:image/png;base64,vue-first');
    const secondCapture = vi.fn(() => ({
      src: 'data:image/png;base64,vue-second',
      alt: 'Vue clip',
    }));
    mounted = await mountVueSpreadsheet({
      props: { captureScreenClip: firstCapture },
    });
    const original = mounted.instance;

    expect(await mounted.instance.captureScreenClip()).toEqual({
      src: 'data:image/png;base64,vue-first',
    });

    await mounted.setProp('captureScreenClip', secondCapture);

    expect(mounted.exposed.instance.value).toBe(original);
    expect(await mounted.instance.captureScreenClip()).toEqual({
      src: 'data:image/png;base64,vue-second',
      alt: 'Vue clip',
    });
    expect(firstCapture).toHaveBeenCalledTimes(1);
    expect(secondCapture).toHaveBeenCalledTimes(1);
  });

  it('keeps the printer profile refresh hook current without remounting', async () => {
    const firstProfile = {
      id: 'vue-first',
      paperSize: 'A4' as const,
      orientation: 'portrait' as const,
      printableBounds: { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 },
    };
    const secondProfile = {
      id: 'vue-second',
      paperSize: 'letter' as const,
      orientation: 'landscape' as const,
      printableBounds: { top: 0.2, right: 0.3, bottom: 0.2, left: 0.3 },
    };
    const firstRefresh = vi.fn(() => [firstProfile]);
    const secondRefresh = vi.fn(() => [secondProfile]);
    mounted = await mountVueSpreadsheet({
      props: { refreshPrinterProfiles: firstRefresh },
    });
    const original = mounted.instance;

    await expect(mounted.instance.refreshPrinterProfiles()).resolves.toEqual([firstProfile]);

    await mounted.setProp('refreshPrinterProfiles', secondRefresh);

    expect(mounted.exposed.instance.value).toBe(original);
    await expect(mounted.instance.refreshPrinterProfiles()).resolves.toEqual([secondProfile]);
    expect(firstRefresh).toHaveBeenCalledTimes(1);
    expect(secondRefresh).toHaveBeenCalledTimes(1);
  });

  it('applies printer profile prop updates to the built-in print flow without remounting', async () => {
    const fallback = {
      id: 'fallback',
      paperSize: 'A4' as const,
      orientation: 'portrait' as const,
      printableBounds: { top: 0.2, right: 0.2, bottom: 0.2, left: 0.2 },
    };
    const selected = {
      id: 'selected',
      paperSize: 'A4' as const,
      orientation: 'portrait' as const,
      printableBounds: { top: 1.2, right: 1.1, bottom: 1.2, left: 1.1 },
    };
    const print = vi.fn();
    const originalPrint = (window as unknown as { print?: () => void }).print;
    (window as unknown as { print: () => void }).print = print;
    try {
      mounted = await mountVueSpreadsheet({
        props: {
          printerProfiles: [fallback, selected],
          printerProfileId: 'fallback',
        },
      });
      const original = mounted.instance;

      await mounted.setProp('printerProfileId', ' selected ');
      mounted.instance.print('print');

      const iframe = mounted.host.querySelector<HTMLIFrameElement>('iframe[data-fc-print-mode]');
      expect(mounted.exposed.instance.value).toBe(original);
      expect(iframe?.srcdoc).toContain(
        '@page { size: A4 portrait; margin: 1.2in 1.1in 1.2in 1.1in; }',
      );
      iframe?.dispatchEvent(new Event('load'));
      iframe?.remove();
    } finally {
      if (originalPrint) (window as unknown as { print: () => void }).print = originalPrint;
      else delete (window as unknown as { print?: () => void }).print;
    }
  });

  it('forwards host status bar prop updates without remounting', async () => {
    mounted = await mountVueSpreadsheet({
      props: {
        uploadStatus: 'saving',
        macroRecording: false,
      },
    });
    const original = mounted.instance;
    mounted.instance.store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        statusOptions: { ...state.ui.statusOptions, uploadStatus: true, macroRecording: true },
      },
    }));

    await mounted.setProp('uploadStatus', 'error');
    await mounted.setProp('macroRecording', true);

    expect(mounted.exposed.instance.value).toBe(original);
    expect(
      mounted.host.querySelector<HTMLElement>('.fc-host__statusbar-upload')?.dataset.uploadStatus,
    ).toBe('error');
    expect(
      mounted.host.querySelector<HTMLElement>('.fc-host__statusbar-macro')?.dataset.macroRecording,
    ).toBe('true');
  });

  it('re-exports Screen Clipping host hook types from the Vue package', async () => {
    const capture: ScreenClipCapture = () => ({
      src: 'data:image/png;base64,vue-export',
      alt: 'Vue export',
    });
    const result = (await capture()) as ScreenClipResult;

    expect(result).toEqual({
      src: 'data:image/png;base64,vue-export',
      alt: 'Vue export',
    });
  });

  it('re-exports shared ribbon and dialog helpers from the Vue package', () => {
    expect(VuePackage.ribbonActivationEntries).toBe(Core.ribbonActivationEntries);
    expect(VuePackage.ribbonSurfaceCommandIds).toBe(Core.ribbonSurfaceCommandIds);
    expect(VuePackage.DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS).toBe(
      Core.DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS,
    );
    expect(VuePackage.attachRangePickerButton).toBe(Core.attachRangePickerButton);
    expect(VuePackage.appendConditionalApplyFormatControls).toBe(
      Core.appendConditionalApplyFormatControls,
    );
    expect(VuePackage.conditionalStyleOptions).toBe(Core.conditionalStyleOptions);
    expect(VuePackage.reportDialogLabels).toBe(Core.reportDialogLabels);
    expect(VuePackage.projectDisabledReason).toBe(Core.projectDisabledReason);
    expect(VuePackage.projectDisabledState).toBe(Core.projectDisabledState);
  });

  it('disposes the engine instance on unmount and unwires event subscriptions', async () => {
    const onCellChange = vi.fn();
    mounted = await mountVueSpreadsheet({ listeners: { onCellChange } });
    const inst = mounted.instance;

    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    await flush();
    expect(onCellChange).toHaveBeenCalledTimes(1);

    await mounted.dispose();
    mounted = null;

    onCellChange.mockClear();
    inst.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    await flush();
    expect(onCellChange).not.toHaveBeenCalled();
  });
});
