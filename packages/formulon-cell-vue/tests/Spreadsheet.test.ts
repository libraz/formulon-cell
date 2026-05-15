import { mutators, WorkbookHandle } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { nextTick } from 'vue';
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

    mutators.setActive(mounted.instance.store, { sheet: 0, row: 4, col: 1 });
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

    const next = await WorkbookHandle.createDefault({ preferStub: true });
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
