import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, h } from 'vue';
import { flush, waitFor } from './test-utils/mount';

const mountCore = vi.hoisted(() => vi.fn());

vi.mock('@libraz/formulon-cell', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@libraz/formulon-cell')>();
  return { ...actual, mount: mountCore };
});

describe('<Spreadsheet> mount failure', () => {
  afterEach(() => {
    mountCore.mockReset();
    document.body.replaceChildren();
  });

  it('emits error and renders errorFallback when the core mount rejects', async () => {
    const err = new Error('engine failed');
    mountCore.mockRejectedValue(err);
    const { Spreadsheet } = await import('../src');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const onError = vi.fn();
    const onReady = vi.fn();

    const app = createApp({
      render: () =>
        h(Spreadsheet, {
          onError,
          onReady,
          errorFallback: (error: unknown) =>
            h('div', { 'data-testid': 'fallback' }, String((error as Error).message)),
        }),
    });
    app.mount(host);
    await waitFor(() => onError.mock.calls.length > 0);
    await flush();

    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(err);
    expect(onReady).not.toHaveBeenCalled();
    expect(host.querySelector('[data-testid="fallback"]')?.textContent).toBe('engine failed');
    app.unmount();
  });

  it('renders no fallback when none is given and the mount rejects', async () => {
    mountCore.mockRejectedValue(new Error('boom'));
    const { Spreadsheet } = await import('../src');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const onError = vi.fn();
    const app = createApp({ render: () => h(Spreadsheet, { onError }) });
    app.mount(host);
    await waitFor(() => onError.mock.calls.length > 0);
    expect(host.querySelector('[data-testid="fallback"]')).toBeNull();
    app.unmount();
  });

  it('passes the host element and only the options that were given to mount', async () => {
    mountCore.mockRejectedValue(new Error('x'));
    const { Spreadsheet } = await import('../src');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const onError = vi.fn();
    const app = createApp({
      render: () => h(Spreadsheet, { onError, locale: 'ja', theme: 'ink' }),
    });
    app.mount(host);
    await waitFor(() => onError.mock.calls.length > 0);
    expect(mountCore).toHaveBeenCalledTimes(1);
    const [el, options] = mountCore.mock.calls[0] as [HTMLElement, Record<string, unknown>];
    expect(el).toBeInstanceOf(HTMLElement);
    expect(options).toEqual({ locale: 'ja', theme: 'ink' });
    app.unmount();
  });
});
