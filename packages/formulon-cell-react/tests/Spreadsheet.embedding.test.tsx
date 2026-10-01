import { act, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mountCore = vi.fn();
const resolveUi = vi.fn(
  (options: { profile?: string; theme?: string; features?: Record<string, boolean> } = {}) => ({
    profile: options.profile ?? 'excel365',
    theme: options.theme ?? 'paper',
    lockTheme: false,
    ribbon: options.features?.ribbon ?? options.profile !== 'embedded',
    print: true,
    features: options.features ?? {},
  }),
);

vi.mock('@libraz/formulon-cell', () => ({
  Spreadsheet: { mount: mountCore },
  resolveSpreadsheetUiOptions: resolveUi,
}));

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
};

const makeInstance = () => ({
  workbook: {},
  i18n: { locale: 'ja', setLocale: vi.fn(), extend: vi.fn() },
  on: vi.fn(() => vi.fn()),
  dispose: vi.fn(),
  setWorkbook: vi.fn(async () => undefined),
  setTheme: vi.fn(),
  setFeatures: vi.fn(),
  setExtensions: vi.fn(),
  setPrinterProfiles: vi.fn(),
  setPrinterProfileId: vi.fn(),
  setUploadStatus: vi.fn(),
  setMacroRecording: vi.fn(),
  setPolicy: vi.fn(),
  setViewportOptions: vi.fn(),
  setContextMenu: vi.fn(),
  setOverlayOptions: vi.fn(),
  setUi: vi.fn(),
  setToolbar: vi.fn(),
});

describe('React Spreadsheet embedding props', () => {
  let root: Root | undefined;
  let host: HTMLDivElement;
  let Spreadsheet: (props: Record<string, unknown>) => ReactNode;

  beforeEach(async () => {
    mountCore.mockReset();
    resolveUi.mockClear();
    ({ Spreadsheet } = await import('../src/Spreadsheet'));
    host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
  });

  afterEach(async () => {
    if (root) {
      await act(async () => {
        root?.unmount();
        await flush();
      });
    }
    root = undefined;
    host.remove();
    vi.restoreAllMocks();
  });

  const render = async (props: Record<string, unknown>): Promise<void> => {
    await act(async () => {
      root?.render(<Spreadsheet {...props} />);
      await flush();
    });
  };

  it('passes embedding options through mount, including an explicit toolbar false', async () => {
    mountCore.mockResolvedValue(makeInstance());
    const policy = { readOnly: true };
    const viewport = { tabBoundary: 'leave' };
    const contextMenu = { mode: 'disabled' };
    const overlays = { root: document.createElement('section') };
    await render({ policy, viewport, contextMenu, overlays, toolbar: false });

    expect(mountCore).toHaveBeenCalledWith(
      expect.any(HTMLElement),
      expect.objectContaining({ policy, viewport, contextMenu, overlays, toolbar: false }),
    );
  });

  it('forwards policy-aware change batches to the latest callback', async () => {
    const instance = makeInstance();
    mountCore.mockResolvedValue(instance);
    const onChangeBatch = vi.fn();
    await render({ onChangeBatch });

    const listener = instance.on.mock.calls.find(([name]) => name === 'changeBatch')?.[1];
    const result = { status: 'applied', applied: [], rejected: [], revision: 1 };
    listener?.(result);

    expect(onChangeBatch).toHaveBeenCalledWith(result);
  });

  it('forwards runtime option removal as undefined and keeps explicit toolbar precedence', async () => {
    const instance = makeInstance();
    mountCore.mockResolvedValue(instance);
    const policy = { readOnly: true };
    const viewport = { tabBoundary: 'stop' };
    const contextMenu = { mode: 'disabled' };
    const overlays = { root: document.createElement('section') };
    await render({
      policy,
      viewport,
      contextMenu,
      overlays,
      ui: { profile: 'embedded' },
      toolbar: true,
    });
    instance.setPolicy.mockClear();
    instance.setViewportOptions.mockClear();
    instance.setContextMenu.mockClear();
    instance.setOverlayOptions.mockClear();
    instance.setUi.mockClear();
    instance.setToolbar.mockClear();

    await render({});

    expect(instance.setPolicy).toHaveBeenCalledWith(undefined);
    expect(instance.setViewportOptions).toHaveBeenCalledWith(undefined);
    expect(instance.setContextMenu).toHaveBeenCalledWith(undefined);
    expect(instance.setOverlayOptions).toHaveBeenCalledWith(undefined);
    expect(instance.setUi).toHaveBeenCalledWith(undefined);
    expect(instance.setToolbar).toHaveBeenCalledWith(false);
  });

  it('reconciles props changed while the async core mount is pending', async () => {
    const instance = makeInstance();
    let resolveMount: ((value: typeof instance) => void) | undefined;
    mountCore.mockImplementation(
      () =>
        new Promise<typeof instance>((resolve) => {
          resolveMount = resolve;
        }),
    );
    const initialPolicy = { readOnly: true };
    const latestPolicy = { defaultOperation: 'allow' };
    await render({ policy: initialPolicy, toolbar: true });
    expect(mountCore).toHaveBeenCalledWith(
      expect.any(HTMLElement),
      expect.objectContaining({ policy: initialPolicy, toolbar: true }),
    );

    await render({ policy: latestPolicy, toolbar: false });
    await act(async () => {
      resolveMount?.(instance);
      await flush();
    });

    expect(instance.setPolicy).toHaveBeenCalledWith(latestPolicy);
    expect(instance.setToolbar).toHaveBeenCalledWith(false);
  });

  it('uses the ui ribbon default only when toolbar is omitted', async () => {
    const instance = makeInstance();
    mountCore.mockResolvedValue(instance);
    await render({ ui: { profile: 'standard' } });
    instance.setToolbar.mockClear();

    await render({ ui: { profile: 'standard' }, toolbar: false });
    expect(instance.setToolbar).toHaveBeenLastCalledWith(false);

    await render({ ui: { profile: 'standard' }, toolbar: true });
    expect(instance.setToolbar).toHaveBeenLastCalledWith(true);
  });
});
