import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';

const mountCore = vi.fn();

vi.mock('@libraz/formulon-cell', () => ({
  Spreadsheet: {
    mount: mountCore,
  },
}));

describe('Vue Spreadsheet error boundary', () => {
  afterEach(() => {
    mountCore.mockReset();
    document.body.replaceChildren();
  });

  it('emits error and renders the custom fallback when core mount fails', async () => {
    const err = new Error('engine failed');
    mountCore.mockImplementation(
      (_host: HTMLElement, opts: { onError?: (error: unknown) => void }) => {
        opts.onError?.(err);
        return Promise.reject(err);
      },
    );
    const { Spreadsheet } = await import('../src/Spreadsheet');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const onError = vi.fn();

    const app = createApp({
      render: () =>
        h(Spreadsheet, {
          onError,
          errorFallback: (error: unknown) =>
            h('div', { 'data-testid': 'fallback' }, String((error as Error).message)),
        }),
    });

    app.mount(host);
    await Promise.resolve();
    await nextTick();

    expect(onError).toHaveBeenCalledWith(err);
    expect(mountCore).toHaveBeenCalledWith(
      expect.any(HTMLElement),
      expect.objectContaining({ renderError: false }),
    );
    expect(host.querySelector('[data-testid="fallback"]')?.textContent).toBe('engine failed');

    app.unmount();
  });
});
