import { type SpreadsheetInstance, Workbook } from '@libraz/formulon-cell';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet, type SpreadsheetProps } from '../src/Spreadsheet';
import { installReactDomStubs, uninstallReactDomStubs } from './test-utils/mount';

const mountSpy = vi.hoisted(() => vi.fn());

vi.mock('@libraz/formulon-cell', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@libraz/formulon-cell')>();
  return { ...actual, mount: mountSpy };
});

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

describe('React Spreadsheet option props', () => {
  let root: Root;
  let host: HTMLDivElement;
  let workbook: Workbook;
  let instance: SpreadsheetInstance | null;
  let gate: Promise<void> | null;

  beforeEach(async () => {
    installReactDomStubs();
    const actual =
      await vi.importActual<typeof import('@libraz/formulon-cell')>('@libraz/formulon-cell');
    gate = null;
    instance = null;
    mountSpy.mockReset();
    mountSpy.mockImplementation(async (el: HTMLElement, options: never) => {
      if (gate) await gate;
      return actual.mount(el, options);
    });
    workbook = await Workbook.createDefault({ engine: 'stub' });
    host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
  });

  afterEach(async () => {
    await act(async () => {
      root.unmount();
      await flush();
    });
    host.remove();
    workbook.dispose();
    uninstallReactDomStubs();
  });

  const render = async (props: SpreadsheetProps): Promise<void> => {
    await act(async () => {
      root.render(
        <Spreadsheet
          workbook={workbook}
          {...props}
          onReady={(i) => {
            instance = i;
            props.onReady?.(i);
          }}
        />,
      );
      await flush();
    });
  };

  const ready = async (): Promise<SpreadsheetInstance> => {
    for (let i = 0; i < 20 && !instance; i += 1) await act(flush);
    if (!instance) throw new Error('the spreadsheet never became ready');
    return instance;
  };

  it('mounts with the option props and leaves the callbacks out of the options', async () => {
    const viewport = { mode: 'fixed', range: 'A1:D20' } as const;
    await render({ viewport, theme: 'ink', locale: 'ja', onChange: () => {} });
    await ready();

    expect(mountSpy).toHaveBeenCalledTimes(1);
    const options = mountSpy.mock.calls[0]?.[1] as Record<string, unknown>;
    expect(options).toMatchObject({ viewport, theme: 'ink', locale: 'ja', workbook });
    expect(Object.keys(options).filter((k) => k.startsWith('on'))).toEqual([]);
    expect(options).not.toHaveProperty('policy');
  });

  it('does not update for structurally equal inline objects', async () => {
    await render({ viewport: { mode: 'fixed', range: 'A1:D20' }, theme: 'ink' });
    const inst = await ready();
    const update = vi.spyOn(inst, 'update');

    await render({ viewport: { mode: 'fixed', range: 'A1:D20' }, theme: 'ink' });
    expect(update).not.toHaveBeenCalled();

    await render({ viewport: { mode: 'fixed', range: 'A1:E30' }, theme: 'ink' });
    expect(update).toHaveBeenCalledTimes(1);
    expect(update).toHaveBeenCalledWith({ viewport: { mode: 'fixed', range: 'A1:E30' } });
  });

  it('resets an option that is no longer set', async () => {
    await render({ theme: 'ink' });
    const inst = await ready();
    const update = vi.spyOn(inst, 'update');

    await render({});

    expect(update).toHaveBeenCalledTimes(1);
    expect(update).toHaveBeenCalledWith({ theme: null });
  });

  it('applies one update per changed key and reports a refused key without blocking the rest', async () => {
    const onError = vi.fn();
    await render({ theme: 'paper', locale: 'en', onError });
    const inst = await ready();
    const refusal = new Error('theme refused');
    const realUpdate = inst.update.bind(inst);
    const update = vi.spyOn(inst, 'update').mockImplementation((patch) => {
      if ('theme' in patch) throw refusal;
      realUpdate(patch);
    });

    await render({ theme: 'ink', locale: 'ja', onError });

    expect(update).toHaveBeenCalledTimes(2);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(refusal);
    expect(inst.i18n.locale).toBe('ja');
  });

  it('reconciles props changed while the mount is pending', async () => {
    let open: () => void = () => {};
    gate = new Promise<void>((resolve) => {
      open = resolve;
    });
    await render({ theme: 'paper', locale: 'en' });
    await render({ theme: 'ink', locale: 'ja' });
    expect(instance).toBeNull();

    await act(async () => {
      open();
      await flush();
    });
    const inst = await ready();

    expect(inst.i18n.locale).toBe('ja');
    expect(host.querySelector<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme).toBe('ink');
    expect(mountSpy).toHaveBeenCalledTimes(1);
  });

  it('disposes an instance that finishes mounting after the component unmounted', async () => {
    let open: () => void = () => {};
    gate = new Promise<void>((resolve) => {
      open = resolve;
    });
    const actual =
      await vi.importActual<typeof import('@libraz/formulon-cell')>('@libraz/formulon-cell');
    let created: SpreadsheetInstance | null = null;
    mountSpy.mockImplementation(async (el: HTMLElement, options: never) => {
      await gate;
      created = await actual.mount(el, options);
      return created;
    });
    const onReady = vi.fn();
    await render({ onReady });

    await act(async () => {
      root.unmount();
      await flush();
    });
    root = createRoot(host);
    await act(async () => {
      open();
      await flush();
    });
    for (let i = 0; i < 20 && !created; i += 1) await act(flush);

    expect(onReady).not.toHaveBeenCalled();
    expect((created as SpreadsheetInstance | null)?.disposed).toBe(true);
  });
});
