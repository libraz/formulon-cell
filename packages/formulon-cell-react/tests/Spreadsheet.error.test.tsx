import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../src/Spreadsheet';

const mountSpy = vi.hoisted(() => vi.fn());

vi.mock('@libraz/formulon-cell', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@libraz/formulon-cell')>();
  return { ...actual, mount: mountSpy };
});

describe('React Spreadsheet error handling', () => {
  let root: Root | null = null;

  afterEach(() => {
    root?.unmount();
    root = null;
    mountSpy.mockReset();
    document.body.replaceChildren();
  });

  const start = (): HTMLDivElement => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
    return host;
  };

  it('reports a rejected mount to onError and renders the function fallback', async () => {
    const err = new Error('engine failed');
    mountSpy.mockRejectedValue(err);
    const host = start();
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

    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(err);
    expect(host.querySelector('[data-testid="fallback"]')?.textContent).toBe('engine failed');
  });

  it('renders a node fallback and renders nothing extra without one', async () => {
    mountSpy.mockRejectedValue(new Error('engine failed'));
    const host = start();

    await act(async () => {
      root?.render(<Spreadsheet errorFallback={<p data-testid="static">failed</p>} />);
      await Promise.resolve();
    });
    expect(host.querySelector('[data-testid="static"]')?.textContent).toBe('failed');

    await act(async () => {
      root?.unmount();
      await Promise.resolve();
    });
    const second = document.createElement('div');
    document.body.appendChild(second);
    root = createRoot(second);
    await act(async () => {
      root?.render(<Spreadsheet />);
      await Promise.resolve();
    });
    expect(second.querySelector('[data-testid="static"]')).toBeNull();
  });

  it('ignores a mount error delivered after the component unmounts', async () => {
    const err = new Error('late engine failure');
    let rejectMount: ((error: unknown) => void) | undefined;
    mountSpy.mockImplementation(
      () =>
        new Promise((_resolve, reject) => {
          rejectMount = reject;
        }),
    );
    start();
    const onError = vi.fn();

    await act(async () => {
      root?.render(<Spreadsheet onError={onError} />);
      await Promise.resolve();
    });
    expect(mountSpy).toHaveBeenCalledTimes(1);

    await act(async () => {
      root?.unmount();
      await Promise.resolve();
    });
    root = null;

    rejectMount?.(err);
    await act(async () => {
      await Promise.resolve();
    });

    expect(onError).not.toHaveBeenCalled();
  });
});
