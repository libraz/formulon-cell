import { WorkbookHandle } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { type MountedReactSpreadsheet, mountReactSpreadsheet } from './test-utils/mount';

/**
 * End-to-end tests for the React `<Spreadsheet>` wrapper. These use the real
 * core (with the stub WASM engine) instead of mocking `Spreadsheet.mount`, so
 * they catch regressions in the prop → instance plumbing in addition to the
 * React layer itself.
 */
describe('React <Spreadsheet>', () => {
  let mounted: MountedReactSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) {
      await mounted.dispose();
      mounted = null;
    }
    document.body.replaceChildren();
  });

  it('invokes onReady with a live SpreadsheetInstance after mount completes', async () => {
    const onReady = vi.fn();
    mounted = await mountReactSpreadsheet({ onReady });

    expect(onReady).toHaveBeenCalledTimes(1);
    const inst = onReady.mock.calls[0]?.[0];
    expect(inst).toBeDefined();
    expect(typeof inst?.dispose).toBe('function');
    expect(typeof inst?.setTheme).toBe('function');
    expect(typeof inst?.on).toBe('function');
    // The forwarded ref should resolve to the same instance.
    expect(mounted.refValue.instance).toBe(inst);
  });

  it('forwards cellChange events from the underlying engine to the onCellChange prop', async () => {
    const onCellChange = vi.fn();
    mounted = await mountReactSpreadsheet({ onCellChange });

    mounted.instance.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 42);

    expect(onCellChange).toHaveBeenCalledTimes(1);
    const event = onCellChange.mock.calls[0]?.[0];
    expect(event?.addr).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(event?.value).toEqual({ kind: 'number', value: 42 });
  });

  it('forwards selectionChange events when the active selection moves', async () => {
    const onSelectionChange = vi.fn();
    mounted = await mountReactSpreadsheet({ onSelectionChange });

    const { mutators } = await import('@libraz/formulon-cell');
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 3, col: 2 });

    expect(onSelectionChange).toHaveBeenCalled();
    const last = onSelectionChange.mock.calls.at(-1)?.[0];
    expect(last?.active).toEqual({ sheet: 0, row: 3, col: 2 });
    expect(last?.range).toMatchObject({ sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 });
  });

  it('forwards workbookChange events when the workbook prop swaps to a new handle', async () => {
    const onWorkbookChange = vi.fn();
    mounted = await mountReactSpreadsheet({ onWorkbookChange });
    const original = mounted.instance.workbook;

    const next = await WorkbookHandle.createDefault({ preferStub: true });
    expect(next).not.toBe(original);
    await mounted.rerender({ workbook: next, onWorkbookChange });

    // setWorkbook is async — wait one extra tick for the inner promise.
    for (let i = 0; i < 10; i += 1) {
      if (onWorkbookChange.mock.calls.length > 0) break;
      await Promise.resolve();
    }

    expect(onWorkbookChange).toHaveBeenCalled();
    const event = onWorkbookChange.mock.calls.at(-1)?.[0];
    expect(event?.workbook).toBe(next);
    expect(mounted.instance.workbook).toBe(next);
  });

  it('forwards a theme prop change to instance.setTheme without remounting', async () => {
    const onThemeChange = vi.fn();
    mounted = await mountReactSpreadsheet({ theme: 'paper', onThemeChange });
    const original = mounted.instance;
    expect(mounted.host.querySelector<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme).toBe(
      'paper',
    );

    await mounted.rerender({ theme: 'ink', onThemeChange });

    expect(mounted.instance).toBe(original); // no remount
    expect(onThemeChange).toHaveBeenCalledTimes(1);
    expect(onThemeChange.mock.calls[0]?.[0]).toEqual({ theme: 'ink' });
    expect(mounted.host.querySelector<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme).toBe('ink');
  });

  it('forwards a locale prop change to instance.i18n.setLocale without remounting', async () => {
    const onLocaleChange = vi.fn();
    mounted = await mountReactSpreadsheet({ locale: 'en', onLocaleChange });
    const original = mounted.instance;
    expect(mounted.instance.i18n.locale).toBe('en');

    await mounted.rerender({ locale: 'ja', onLocaleChange });

    expect(mounted.instance).toBe(original);
    expect(mounted.instance.i18n.locale).toBe('ja');
    expect(onLocaleChange).toHaveBeenCalled();
    const event = onLocaleChange.mock.calls.at(-1)?.[0];
    expect(event?.locale).toBe('ja');
  });

  it('disposes the engine instance on unmount and unwires event subscriptions', async () => {
    const onCellChange = vi.fn();
    mounted = await mountReactSpreadsheet({ onCellChange });
    const inst = mounted.instance;

    // Sanity: events flow before unmount.
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    expect(onCellChange).toHaveBeenCalledTimes(1);

    await mounted.dispose();
    mounted = null;

    // After unmount the React-bound subscription is gone, so the user-supplied
    // callback no longer fires regardless of what the workbook reports back.
    onCellChange.mockClear();
    inst.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    // Drain microtasks so any straggling listeners have a chance to fire.
    for (let i = 0; i < 5; i += 1) await Promise.resolve();
    expect(onCellChange).not.toHaveBeenCalled();
  });
});
