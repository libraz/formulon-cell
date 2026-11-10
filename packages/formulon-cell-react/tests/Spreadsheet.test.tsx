import type * as Core from '@libraz/formulon-cell';
import { Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { type MountedReactSpreadsheet, mountReactSpreadsheet } from './test-utils/mount';

/**
 * End-to-end tests for the React `<Spreadsheet>` component against the real core (stub engine),
 * so they cover the prop to instance plumbing as well as the React layer.
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

  it('invokes onReady with a live SpreadsheetInstance and exposes it through the ref', async () => {
    const onReady = vi.fn();
    mounted = await mountReactSpreadsheet({ onReady });

    expect(onReady).toHaveBeenCalledTimes(1);
    const inst = onReady.mock.calls[0]?.[0] as Core.SpreadsheetInstance;
    expect(inst.disposed).toBe(false);
    expect(typeof inst.update).toBe('function');
    expect(mounted.refValue.spreadsheet).toBe(inst);
  });

  it('defaults the locale to en', async () => {
    mounted = await mountReactSpreadsheet();
    expect(mounted.instance.i18n.locale).toBe('en');
  });

  it('renders a render-prop child once the instance is mounted', async () => {
    const renderChild = vi.fn((instance: Core.SpreadsheetInstance) => (
      <span data-testid="render-prop-child">{instance.i18n.locale}</span>
    ));
    mounted = await mountReactSpreadsheet({ locale: 'ja' }, renderChild as never);

    expect(renderChild).toHaveBeenCalledTimes(1);
    expect(mounted.host.querySelector('[data-testid="render-prop-child"]')?.textContent).toBe('ja');
  });

  it('forwards cell events from the workbook to onCell', async () => {
    const onCell = vi.fn();
    mounted = await mountReactSpreadsheet({ onCell });

    mounted.instance.workbook.sheet(0).setNumber(0, 0, 42);

    expect(onCell).toHaveBeenCalledTimes(1);
    const event = onCell.mock.calls[0]?.[0];
    expect(event).toMatchObject({ sheet: 0, at: { row: 0, col: 0 } });
    expect(event.value).toMatchObject({ number: 42 });
  });

  it('forwards change and recalc events', async () => {
    const onChange = vi.fn();
    const onRecalc = vi.fn();
    mounted = await mountReactSpreadsheet({ onChange, onRecalc });

    const result = mounted.instance.applyChanges([{ at: { row: 0, col: 0 }, input: '7' }], {
      history: 'record',
    });

    expect(result.applied.length).toBeGreaterThan(0);
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0]?.[0].origin).toBe('instanceApi');
    expect(onRecalc).toHaveBeenCalled();
  });

  it('forwards selection changes to onSelectionChange', async () => {
    const onSelectionChange = vi.fn();
    mounted = await mountReactSpreadsheet({ onSelectionChange });

    mounted.instance.selection.set('C4');

    expect(onSelectionChange).toHaveBeenCalled();
    const last = onSelectionChange.mock.calls.at(-1)?.[0];
    expect(last.active).toMatchObject({ row: 3, col: 2 });
  });

  it('forwards a workbook swap to onWorkbook and the instance', async () => {
    const onWorkbook = vi.fn();
    mounted = await mountReactSpreadsheet({ onWorkbook });
    const next = await Workbook.createDefault({ engine: 'stub' });

    try {
      await mounted.rerender({ workbook: next, onWorkbook });
      for (let i = 0; i < 10 && onWorkbook.mock.calls.length === 0; i += 1) await Promise.resolve();

      expect(onWorkbook).toHaveBeenCalled();
      expect(onWorkbook.mock.calls.at(-1)?.[0].workbook).toBe(next);
      expect(mounted.instance.workbook).toBe(next);
    } finally {
      await mounted.rerender({ workbook: mounted.workbook });
      next.dispose();
    }
  });

  it('applies a theme prop change with update and reports it through onUpdate', async () => {
    const onUpdate = vi.fn();
    mounted = await mountReactSpreadsheet({ theme: 'paper', onUpdate });
    const original = mounted.instance;
    expect(mounted.host.querySelector<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme).toBe(
      'paper',
    );

    await mounted.rerender({ theme: 'ink', onUpdate });

    expect(mounted.refValue.spreadsheet).toBe(original);
    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(onUpdate.mock.calls[0]?.[0]).toEqual({ keys: ['theme'] });
    expect(mounted.host.querySelector<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme).toBe('ink');
  });

  it('applies a locale prop change and reports it through onLocaleChange', async () => {
    const onLocaleChange = vi.fn();
    mounted = await mountReactSpreadsheet({ locale: 'en', onLocaleChange });
    const original = mounted.instance;

    await mounted.rerender({ locale: 'ja', onLocaleChange });

    expect(mounted.refValue.spreadsheet).toBe(original);
    expect(original.i18n.locale).toBe('ja');
    expect(onLocaleChange).toHaveBeenCalled();
    expect(onLocaleChange.mock.calls.at(-1)?.[0].locale).toBe('ja');
  });

  it('keeps the printer profile refresh provider current without remounting', async () => {
    const first = vi.fn(() => []);
    const second = vi.fn(() => []);
    mounted = await mountReactSpreadsheet({ host: { refreshPrinterProfiles: first } });
    const original = mounted.instance;

    await original.refreshPrinterProfiles();
    const firstCalls = first.mock.calls.length;
    expect(firstCalls).toBeGreaterThan(0);

    await mounted.rerender({ host: { refreshPrinterProfiles: second } });
    second.mockClear();
    await original.refreshPrinterProfiles();

    expect(mounted.refValue.spreadsheet).toBe(original);
    expect(second).toHaveBeenCalledTimes(1);
    expect(first).toHaveBeenCalledTimes(firstCalls);
  });

  it('keeps the Screen Clipping capture provider current without remounting', async () => {
    const first = vi.fn(() => 'data:image/png;base64,first');
    const second = vi.fn(() => ({ src: 'data:image/png;base64,second', alt: 'second clip' }));
    mounted = await mountReactSpreadsheet({ host: { captureScreenClip: first } });
    const original = mounted.instance;

    expect(await original.captureScreenClip()).toEqual({ src: 'data:image/png;base64,first' });
    await mounted.rerender({ host: { captureScreenClip: second } });

    expect(mounted.refValue.spreadsheet).toBe(original);
    expect(await original.captureScreenClip()).toEqual({
      src: 'data:image/png;base64,second',
      alt: 'second clip',
    });
    expect(first).toHaveBeenCalledTimes(1);
    expect(second).toHaveBeenCalledTimes(1);
  });

  it('keeps the function argument help provider current without remounting', async () => {
    const first = vi.fn(() => ({ description: 'first hint' }));
    const second = vi.fn((_name: string, _index: number, locale: string) => ({
      description: `second hint (${locale})`,
    }));
    const common = { ui: { platform: 'mac' }, locale: 'en' } as const;
    mounted = await mountReactSpreadsheet({ ...common, host: { functionArgumentHelp: first } });
    const original = mounted.instance;
    const hint = (): string | null | undefined =>
      document.querySelector('.fc-mac-formula-palette__argument small')?.textContent;
    const closePalette = (): void =>
      document
        .querySelector<HTMLButtonElement>('.fc-mac-formula-palette [data-action="close"]')
        ?.click();

    original.dialogs.open('functionArguments', { name: 'ACOS' });
    expect(hint()).toBe('first hint');
    closePalette();

    await mounted.rerender({ ...common, host: { functionArgumentHelp: second } });
    original.dialogs.open('functionArguments', { name: 'ACOS' });

    expect(mounted.refValue.spreadsheet).toBe(original);
    expect(hint()).toBe('second hint (en)');
    expect(second).toHaveBeenCalledWith('ACOS', 0, 'en');
    closePalette();
  });

  it('clears the host providers when the host prop is removed', async () => {
    const capture = vi.fn(() => 'data:image/png;base64,clip');
    const refresh = vi.fn(() => []);
    mounted = await mountReactSpreadsheet({
      host: { captureScreenClip: capture, refreshPrinterProfiles: refresh },
    });
    const original = mounted.instance;
    expect(await original.captureScreenClip()).toEqual({ src: 'data:image/png;base64,clip' });

    await mounted.rerender({});
    const refreshCalls = refresh.mock.calls.length;
    await original.refreshPrinterProfiles();

    expect(mounted.refValue.spreadsheet).toBe(original);
    expect(await original.captureScreenClip()).toBeNull();
    expect(capture).toHaveBeenCalledTimes(1);
    expect(refresh).toHaveBeenCalledTimes(refreshCalls);
  });

  it('disposes the instance on unmount and clears the ref', async () => {
    const onCell = vi.fn();
    mounted = await mountReactSpreadsheet({ onCell });
    const inst = mounted.instance;
    const refValue = mounted.refValue;
    inst.workbook.sheet(0).setNumber(0, 0, 1);
    expect(onCell).toHaveBeenCalledTimes(1);

    await mounted.dispose();
    mounted = null;

    expect(inst.disposed).toBe(true);
    expect(refValue.spreadsheet).toBeNull();
  });
});
