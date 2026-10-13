import { act, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mountCore = vi.fn();
const resolveUi = vi.fn(
  (options: { profile?: string; theme?: string; features?: Record<string, boolean> } = {}) => ({
    profile: options.profile ?? 'excel365',
    theme: options.theme ?? 'paper',
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
  locale: 'ja',
  i18n: { locale: 'ja', strings: {}, setLocale: vi.fn(), extend: vi.fn() },
  setLocale: vi.fn(() => true),
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

  it('keeps the original workbook usable after a pending controlled replacement rejects', async () => {
    const original = {
      getValue: vi.fn(() => ({ kind: 'text', value: 'original' })),
      dispose: vi.fn(),
    };
    const next = { getValue: vi.fn(() => ({ kind: 'text', value: 'next' })), dispose: vi.fn() };
    const instance = makeInstance();
    instance.workbook = original;
    const failure = new Error('controlled workbook rejected');
    instance.setWorkbook.mockRejectedValueOnce(failure);
    let resolveMount: ((value: typeof instance) => void) | undefined;
    mountCore.mockImplementation(
      () =>
        new Promise<typeof instance>((resolve) => {
          resolveMount = resolve;
        }),
    );
    const onError = vi.fn();

    await render({ workbook: original, onError });
    await render({ workbook: next, onError });
    await act(async () => {
      resolveMount?.(instance);
      await flush();
    });

    expect(instance.setWorkbook).toHaveBeenCalledTimes(1);
    expect(instance.setWorkbook).toHaveBeenCalledWith(next);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(failure);
    expect(instance.workbook).toBe(original);
    expect(original.getValue()).toEqual({ kind: 'text', value: 'original' });
    expect(original.dispose).not.toHaveBeenCalled();
    expect(next.dispose).not.toHaveBeenCalled();

    await act(async () => {
      root?.unmount();
      await flush();
    });
    root = undefined;
    expect(original.dispose).not.toHaveBeenCalled();
    expect(next.dispose).not.toHaveBeenCalled();
  });

  it('continues pending controlled prop reconciliation after workbook rejection', async () => {
    const original = {
      getValue: vi.fn(() => ({ kind: 'text', value: 'original' })),
      dispose: vi.fn(),
    };
    const next = { getValue: vi.fn(() => ({ kind: 'text', value: 'next' })), dispose: vi.fn() };
    const instance = makeInstance();
    instance.workbook = original;
    const failure = new Error('controlled workbook rejected');
    instance.setWorkbook.mockRejectedValueOnce(failure);
    instance.setLocale.mockImplementation((nextLocale) => {
      instance.locale = nextLocale;
      instance.i18n.locale = nextLocale;
      return true;
    });
    let resolveMount: ((value: typeof instance) => void) | undefined;
    mountCore.mockImplementation(
      () =>
        new Promise<typeof instance>((resolve) => {
          resolveMount = resolve;
        }),
    );
    const onError = vi.fn();
    const initialPolicy = { readOnly: true };
    const latestPolicy = { defaultOperation: 'allow' };
    const strings = { ribbon: { bold: 'Stable bold' } };
    const features = { ribbon: false };

    await render({ workbook: original, policy: initialPolicy, locale: 'ja', onError });
    await render({
      workbook: next,
      policy: latestPolicy,
      locale: 'en',
      strings,
      features,
      onError,
    });
    await act(async () => {
      resolveMount?.(instance);
      await flush();
    });

    expect(instance.setWorkbook).toHaveBeenCalledTimes(1);
    expect(instance.setWorkbook).toHaveBeenCalledWith(next);
    expect(instance.setPolicy).toHaveBeenCalledWith(latestPolicy);
    expect(instance.setLocale).toHaveBeenCalledWith('en');
    expect(instance.i18n.extend).toHaveBeenCalledWith('en', strings);
    expect(instance.setFeatures).toHaveBeenCalledWith(expect.objectContaining(features));
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(failure);
    expect(instance.workbook).toBe(original);
    expect(original.getValue()).toEqual({ kind: 'text', value: 'original' });
    expect(original.dispose).not.toHaveBeenCalled();
    expect(next.dispose).not.toHaveBeenCalled();
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

  it('uses the public locale setter and reconciles a rejected controlled update', async () => {
    const instance = makeInstance();
    mountCore.mockResolvedValue(instance);
    const onLocaleChange = vi.fn();
    const onError = vi.fn();

    await render({ locale: 'en', onLocaleChange, onError });
    instance.setLocale.mockReturnValue(false);
    await render({ locale: 'fr', onLocaleChange, onError });

    expect(instance.setLocale).toHaveBeenCalledWith('fr');
    expect(onLocaleChange).toHaveBeenCalledWith({ locale: 'ja', strings: instance.i18n.strings });
    expect(onError).toHaveBeenCalledWith(expect.any(Error));
  });

  it('delivers a rollback event after a prior successful switch to the same locale', async () => {
    const instance = makeInstance();
    mountCore.mockResolvedValue(instance);
    let localeHandler: ((event: { locale: string; strings: object }) => void) | undefined;
    instance.on.mockImplementation((name, handler) => {
      if (name === 'localeChange') {
        localeHandler = handler as (event: { locale: string; strings: object }) => void;
      }
      return vi.fn();
    });
    instance.setLocale.mockImplementation((next) => {
      if (next === 'ja') {
        localeHandler?.({ locale: 'ja', strings: instance.i18n.strings });
      }
      return next !== 'fr';
    });
    const onLocaleChange = vi.fn();

    await render({ locale: 'en', onLocaleChange });
    onLocaleChange.mockClear();
    await render({ locale: 'ja', onLocaleChange });
    expect(onLocaleChange).toHaveBeenCalledTimes(1);

    onLocaleChange.mockClear();
    await render({ locale: 'fr', onLocaleChange });

    expect(onLocaleChange).toHaveBeenCalledTimes(1);
    expect(onLocaleChange).toHaveBeenCalledWith({
      locale: 'ja',
      strings: instance.i18n.strings,
    });
  });

  it('reapplies a stable strings overlay after each accepted locale switch', async () => {
    const instance = makeInstance();
    mountCore.mockResolvedValue(instance);
    const strings = { ribbon: { bold: 'Stable bold' } };
    instance.setLocale.mockImplementation((next) => {
      instance.locale = next;
      instance.i18n.locale = next;
      return true;
    });

    await render({ locale: 'en', strings });
    instance.setLocale.mockClear();
    instance.i18n.extend.mockClear();

    for (const locale of ['ja', 'de', 'fr', 'zh', 'ko', 'th', 'en']) {
      await render({ locale, strings });
    }

    expect(instance.setLocale.mock.calls.map(([locale]) => locale)).toEqual([
      'ja',
      'de',
      'fr',
      'zh',
      'ko',
      'th',
      'en',
    ]);
    expect(instance.i18n.extend.mock.calls).toEqual(
      ['ja', 'de', 'fr', 'zh', 'ko', 'th', 'en'].map((locale) => [locale, strings]),
    );
  });

  it('does not reapply a stable strings overlay after a rejected locale switch', async () => {
    const instance = makeInstance();
    mountCore.mockResolvedValue(instance);
    const strings = { ribbon: { bold: 'Stable bold' } };
    const failure = new Error('locale rejected');
    const onError = vi.fn();
    instance.setLocale.mockImplementation((next) => {
      if (next === 'fr') return false;
      instance.locale = next;
      instance.i18n.locale = next;
      return true;
    });

    await render({ locale: 'en', strings, onError });
    instance.i18n.extend.mockClear();
    instance.setLocale.mockImplementationOnce(() => {
      throw failure;
    });

    await render({ locale: 'fr', strings, onError });

    expect(instance.i18n.extend).not.toHaveBeenCalled();
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(failure);
  });
});
