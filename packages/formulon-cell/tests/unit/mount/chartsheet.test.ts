import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { setSheetZoom } from '../../../src/commands/view.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { zipEntries, zipText } from '../engine/zip-entries.js';

const fixture = (): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, '../../fixtures/charts/chart_chartsheet.xlsx')));

describe('an active chartsheet', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  const mountFixture = async (): Promise<MountedStubSheet> => {
    const workbook = await WorkbookHandle.loadBytes(fixture());
    sheet = await mountStubSheet({ workbook });
    return sheet;
  };

  it('marks the host and draws its chart in place of the grid', async () => {
    const { host, instance } = await mountFixture();
    expect(instance.store.getState().data.sheetIndex).toBe(0);
    expect(host.dataset.fcSheetKind).toBe('chartsheet');
    const panels = host.querySelectorAll('.fc-charts .fc-chart');
    expect(panels).toHaveLength(1);
    // Its chart is fixed to the sheet: no sizing handles.
    expect(panels[0]?.classList.contains('fc-chart--fixed')).toBe(true);
  });

  it('draws the chart over the whole area of the grid', async () => {
    const { host, instance } = await mountFixture();
    const grid = host.querySelector<HTMLElement>('.fc-host__grid');
    if (!grid) throw new Error('missing grid');
    Object.defineProperty(grid, 'clientWidth', { value: 800, configurable: true });
    Object.defineProperty(grid, 'clientHeight', { value: 500, configurable: true });
    mutators.setSheetIndex(instance.store, 1);
    mutators.setSheetIndex(instance.store, 0);
    const panel = host.querySelector<HTMLElement>('.fc-charts .fc-chart');
    expect(panel?.style.width).toBe('800px');
    expect(panel?.style.height).toBe('500px');
    expect(panel?.style.left).toBe('0px');
    expect(panel?.style.top).toBe('0px');
  });

  it('shows nothing editable in the name box and the formula bar', async () => {
    const { host } = await mountFixture();
    const tag = host.querySelector<HTMLInputElement>('.fc-host__formulabar-tag');
    const fx = host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    expect(tag?.value).toBe('');
    expect(tag?.readOnly).toBe(true);
    expect(fx?.value).toBe('');
    expect(fx?.readOnly).toBe(true);
  });

  it('does not open the cell editor', async () => {
    const { instance } = await mountFixture();
    instance.cellEditor.begin('x');
    expect(instance.cellEditor.isActive()).toBe(false);
    mutators.setSheetIndex(instance.store, 1);
    instance.cellEditor.begin('x');
    expect(instance.cellEditor.isActive()).toBe(true);
  });

  it('saves the fit-to-window zoom as the file states it', async () => {
    const { instance } = await mountFixture();
    mutators.setSheetIndex(instance.store, 1);
    mutators.setSheetIndex(instance.store, 0);
    const entries = zipEntries(instance.workbook.save());
    const path = [...entries.keys()].find((key) => key.startsWith('xl/chartsheets/sheet'));
    const view = (path ? zipText(entries, path) : '').match(/<sheetView\b[^>]*>/)?.[0] ?? '';
    expect(view).toContain('zoomScale="125"');
    expect(view).toContain('zoomToFit="1"');
  });

  describe('zoom', () => {
    const savedView = (workbook: WorkbookHandle): string => {
      const entries = zipEntries(workbook.save());
      const path = [...entries.keys()].find((key) => key.startsWith('xl/chartsheets/sheet'));
      return (path ? zipText(entries, path) : '').match(/<sheetView\b[^>]*>/)?.[0] ?? '';
    };
    const wheelZoom = (host: HTMLElement, deltaY: number): void => {
      const grid = host.querySelector<HTMLElement>('.fc-host__grid');
      const event = new WheelEvent('wheel', { deltaY, bubbles: true, cancelable: true });
      // happy-dom drops modifier keys from the WheelEvent init bag.
      Object.defineProperty(event, 'ctrlKey', { value: true });
      grid?.dispatchEvent(event);
    };

    it('saves a zoom set on it as that zoom, ending the fit to the window', async () => {
      const { instance, workbook } = await mountFixture();
      setSheetZoom(instance.store, 0.8, workbook);
      expect(instance.store.getState().viewport.zoom).toBe(0.8);
      const view = savedView(workbook);
      expect(view).toContain('zoomScale="80"');
      expect(view).not.toContain('zoomToFit');
    });

    it('zooms through the wheel and the status bar', async () => {
      const { host, instance, workbook } = await mountFixture();
      // Kept clear of 100%, the default a saved view leaves unwritten.
      setSheetZoom(instance.store, 0.5, workbook);
      wheelZoom(host, -100);
      const wheeled = instance.store.getState().viewport.zoom;
      expect(wheeled).toBeGreaterThan(0.5);
      host.querySelector<HTMLButtonElement>('.fc-host__statusbar-zoom-btn')?.click();
      expect(instance.store.getState().viewport.zoom).toBeLessThan(wheeled);
      const pct = Math.round(instance.store.getState().viewport.zoom * 100);
      const view = savedView(workbook);
      expect(view).toContain(`zoomScale="${pct}"`);
      expect(view).not.toContain('zoomToFit');
    });

    it('runs the zoom commands but refuses zoom to selection, which needs cells', async () => {
      const { instance } = await mountFixture();
      for (const id of ['zoom100', 'zoom125', 'zoom75', 'zoomDialog', 'mac.view.zoom']) {
        expect(canExecuteBuiltIn(instance.store, id as never, 'ribbon').allowed, id).toBe(true);
      }
      for (const id of ['zoomSelection', 'mac.view.zoomSelection']) {
        expect(canExecuteBuiltIn(instance.store, id as never, 'ribbon'), id).toMatchObject({
          allowed: false,
          reasonKey: 'sheetHasNoCells',
        });
      }
    });

    it('still zooms a worksheet', async () => {
      const { host, instance, workbook } = await mountFixture();
      mutators.setSheetIndex(instance.store, 1);
      expect(canExecuteBuiltIn(instance.store, 'zoom125' as never, 'ribbon').allowed).toBe(true);
      setSheetZoom(instance.store, 1.5, workbook);
      expect(instance.store.getState().viewport.zoom).toBe(1.5);
      wheelZoom(host, -100);
      expect(instance.store.getState().viewport.zoom).toBeGreaterThan(1.5);
      host.querySelector<HTMLButtonElement>('.fc-host__statusbar-zoom-btn')?.click();
      expect(instance.store.getState().viewport.zoom).toBeLessThan(1.7);
    });
  });

  it('is a grid again on the next worksheet tab', async () => {
    const { host, instance } = await mountFixture();
    mutators.setSheetIndex(instance.store, 1);
    expect(host.dataset.fcSheetKind).toBeUndefined();
    const tag = host.querySelector<HTMLInputElement>('.fc-host__formulabar-tag');
    expect(tag?.readOnly).toBe(false);
  });
});
