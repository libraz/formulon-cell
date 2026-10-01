import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

const mountCore = vi.fn();

vi.mock('@libraz/formulon-cell', () => ({
  Spreadsheet: {
    mount: mountCore,
  },
}));

describe('React Spreadsheet error boundary', () => {
  let root: Root | null = null;

  afterEach(() => {
    root?.unmount();
    root = null;
    mountCore.mockReset();
    document.body.replaceChildren();
  });

  it('emits onError and renders the custom fallback when core mount fails', async () => {
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
    root = createRoot(host);
    const onError = vi.fn();

    await act(async () => {
      root?.render(
        <Spreadsheet
          onError={onError}
          errorFallback={(error) => (
            <div data-testid="fallback">{String((error as Error).message)}</div>
          )}
        />,
      );
      await Promise.resolve();
    });

    expect(onError).toHaveBeenCalledWith(err);
    expect(mountCore).toHaveBeenCalledWith(
      expect.any(HTMLElement),
      expect.objectContaining({ renderError: false }),
    );
    expect(host.querySelector('[data-testid="fallback"]')?.textContent).toBe('engine failed');
  });

  it('ignores a mount error delivered after the component unmounts', async () => {
    const err = new Error('late engine failure');
    let reportError: ((error: unknown) => void) | undefined;
    let rejectMount: ((error: unknown) => void) | undefined;
    mountCore.mockImplementation(
      (_host: HTMLElement, opts: { onError?: (error: unknown) => void }) => {
        reportError = opts.onError;
        return new Promise((_resolve, reject) => {
          rejectMount = reject;
        });
      },
    );
    const { Spreadsheet } = await import('../src/Spreadsheet');
    const host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
    const onError = vi.fn();

    await act(async () => {
      root?.render(<Spreadsheet onError={onError} />);
      await Promise.resolve();
    });
    expect(mountCore).toHaveBeenCalledTimes(1);

    await act(async () => {
      root?.unmount();
      await Promise.resolve();
    });
    root = null;

    reportError?.(err);
    rejectMount?.(err);
    await act(async () => {
      await Promise.resolve();
    });

    expect(onError).not.toHaveBeenCalled();
  });
});
