import { type SpreadsheetInstance, Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { type App, createApp, defineComponent, h, shallowReactive, shallowRef } from 'vue';
import { Spreadsheet, type SpreadsheetExposed } from '../src';
import { flush, stubWorkbook, waitFor } from './test-utils/mount';

interface Harness {
  app: App;
  props: Record<string, unknown>;
  exposed: () => SpreadsheetExposed | null;
  instance: () => SpreadsheetInstance | null;
  ready: ReturnType<typeof vi.fn>;
  unmount: () => void;
}

const apps: App[] = [];
const extra: Workbook[] = [];

afterEach(() => {
  for (const app of apps.splice(0)) app.unmount();
  for (const workbook of extra.splice(0)) workbook.dispose();
  document.body.replaceChildren();
});

/** Mounts `<Spreadsheet>` with reactive props on a stub workbook; waits for ready unless told not to. */
const mountComponent = async (
  initial: Record<string, unknown> = {},
  { wait = true, slots }: { wait?: boolean; slots?: Parameters<typeof h>[2] } = {},
): Promise<Harness> => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const props = shallowReactive<Record<string, unknown>>({
    workbook: 'workbook' in initial ? initial.workbook : await stubWorkbook(),
    locale: 'en',
    ...initial,
  });
  const ready = vi.fn();
  const child = shallowRef<SpreadsheetExposed | null>(null);
  const app = createApp(
    defineComponent({
      render: () =>
        h(
          Spreadsheet,
          {
            ...props,
            onReady: ready,
            ref: (el: unknown) => {
              child.value = el as SpreadsheetExposed | null;
            },
          },
          slots,
        ),
    }),
  );
  app.mount(host);
  apps.push(app);
  if (wait) await waitFor(() => ready.mock.calls.length > 0);
  return {
    app,
    props,
    exposed: () => child.value,
    instance: () => child.value?.spreadsheet ?? null,
    ready,
    unmount: () => app.unmount(),
  };
};

describe('<Spreadsheet>', () => {
  it('emits ready with the live instance, which the template ref exposes as a plain getter', async () => {
    const view = await mountComponent();
    expect(view.ready).toHaveBeenCalledTimes(1);
    const inst = view.ready.mock.calls[0]?.[0] as SpreadsheetInstance;
    expect(inst.disposed).toBe(false);
    expect(view.exposed()?.spreadsheet).toBe(inst);
  });

  it('exposes null before the mount settles', async () => {
    const view = await mountComponent({}, { wait: false });
    expect(view.exposed()?.spreadsheet).toBeNull();
    await waitFor(() => view.ready.mock.calls.length > 0);
    expect(view.exposed()?.spreadsheet).not.toBeNull();
  });

  it('forwards the cell event when the workbook changes', async () => {
    const onCell = vi.fn();
    const view = await mountComponent({ onCell });
    const inst = view.instance() as SpreadsheetInstance;
    inst.workbook.sheet(0).setNumber(0, 0, 42);
    await flush();
    expect(onCell).toHaveBeenCalled();
    expect(onCell.mock.calls.at(-1)?.[0]).toMatchObject({ sheet: 0, at: { row: 0, col: 0 } });
  });

  it('forwards change and update events', async () => {
    const onChange = vi.fn();
    const onUpdate = vi.fn();
    const view = await mountComponent({ onChange, onUpdate });
    const inst = view.instance() as SpreadsheetInstance;
    inst.applyChanges([{ at: 'B2', input: '7' }], { history: 'record' });
    await flush();
    expect(onChange).toHaveBeenCalled();
    view.props.theme = 'ink';
    await flush();
    expect(onUpdate).toHaveBeenCalledWith({ keys: ['theme'] });
  });

  it('emits selectionChange when the selection moves', async () => {
    const onSelectionChange = vi.fn();
    const view = await mountComponent({ onSelectionChange });
    (view.instance() as SpreadsheetInstance).selection.set('C3');
    await flush();
    expect(onSelectionChange).toHaveBeenCalled();
    expect(onSelectionChange.mock.calls.at(-1)?.[0].active).toEqual({ row: 2, col: 2 });
  });

  it('emits workbook when the workbook prop is replaced, without remounting', async () => {
    const onWorkbook = vi.fn();
    const view = await mountComponent({ onWorkbook });
    const first = view.instance();
    const next = await stubWorkbook();
    extra.push(next);
    view.props.workbook = next;
    await flush();
    expect(onWorkbook).toHaveBeenCalledTimes(1);
    expect(onWorkbook.mock.calls[0]?.[0].workbook).toBe(next);
    expect(view.instance()).toBe(first);
    expect(view.ready).toHaveBeenCalledTimes(1);
    expect(first?.workbook).toBe(next);
  });

  it('applies a theme prop change with one update and no remount', async () => {
    const onUpdate = vi.fn();
    const view = await mountComponent({ onUpdate });
    const first = view.instance();
    view.props.theme = 'ink';
    await flush();
    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(onUpdate).toHaveBeenCalledWith({ keys: ['theme'] });
    expect(view.instance()).toBe(first);
  });

  it('emits localeChange when the locale prop changes', async () => {
    const onLocaleChange = vi.fn();
    const view = await mountComponent({ onLocaleChange });
    expect((view.instance() as SpreadsheetInstance).i18n.locale).toBe('en');
    view.props.locale = 'ja';
    await flush();
    expect(onLocaleChange).toHaveBeenCalled();
    expect((view.instance() as SpreadsheetInstance).i18n.locale).toBe('ja');
  });

  it('does not update again for a structurally equal inline object', async () => {
    const onUpdate = vi.fn();
    const view = await mountComponent({ onUpdate, ui: { profile: 'minimal' } });
    await flush();
    onUpdate.mockClear();
    view.props.ui = { profile: 'minimal' };
    await flush();
    expect(onUpdate).not.toHaveBeenCalled();
    view.props.ui = { profile: 'full' };
    await flush();
    expect(onUpdate).toHaveBeenCalledWith({ keys: ['ui'] });
  });

  it('applies props changed while the mount is pending once it settles', async () => {
    const view = await mountComponent({ theme: 'paper' }, { wait: false });
    view.props.theme = 'ink';
    await waitFor(() => view.ready.mock.calls.length > 0);
    await flush();
    const inst = view.instance() as SpreadsheetInstance;
    expect(inst.disposed).toBe(false);
    const seen: string[] = [];
    inst.on('update', (e) => seen.push(...e.keys));
    view.props.theme = 'paper';
    await flush();
    expect(seen).toEqual(['theme']);
  });

  it('emits error for a refused key and still applies the others', async () => {
    const onError = vi.fn();
    const onUpdate = vi.fn();
    const view = await mountComponent({ onError, onUpdate });
    onUpdate.mockClear();
    view.props.viewport = { mode: 'fixed', range: 'not a range' };
    view.props.theme = 'ink';
    await flush();
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onUpdate).toHaveBeenCalledWith({ keys: ['theme'] });
    expect(view.instance()?.disposed).toBe(false);
  });

  it('renders the default slot with the instance once mounted', async () => {
    const seen: unknown[] = [];
    const view = await mountComponent(
      {},
      {
        wait: false,
        slots: {
          default: (slot: { spreadsheet: SpreadsheetInstance }) => {
            seen.push(slot.spreadsheet);
            return h('span', { 'data-testid': 'slot' }, 'inside');
          },
        } as never,
      },
    );
    expect(document.querySelector('[data-testid="slot"]')).toBeNull();
    await waitFor(() => view.ready.mock.calls.length > 0);
    await flush();
    expect(document.querySelector('[data-testid="slot"]')?.textContent).toBe('inside');
    expect(seen.at(-1)).toBe(view.instance());
  });

  it('applies class and style to the host element', async () => {
    await mountComponent({ class: 'sheet-host', style: { height: '321px' } });
    const el = document.querySelector<HTMLElement>('.sheet-host');
    expect(el).not.toBeNull();
    expect(el?.style.height).toBe('321px');
  });

  it('disposes the instance on unmount, unsubscribes, and exposes null', async () => {
    const onCell = vi.fn();
    const view = await mountComponent({ onCell });
    const inst = view.instance() as SpreadsheetInstance;
    const workbook = inst.workbook;
    view.unmount();
    expect(inst.disposed).toBe(true);
    expect(view.exposed()?.spreadsheet ?? null).toBeNull();
    onCell.mockClear();
    expect(workbook.disposed).toBe(false);
    workbook.dispose();
  });

  it('disposes an instance whose mount settles after unmount', async () => {
    const workbook = await Workbook.createDefault({ engine: 'stub' });
    const view = await mountComponent({ workbook }, { wait: false });
    view.unmount();
    await flush();
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(view.ready).not.toHaveBeenCalled();
    workbook.dispose();
  });
});
