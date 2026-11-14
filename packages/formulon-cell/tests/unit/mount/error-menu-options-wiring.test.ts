import { afterEach, describe, expect, it, vi } from 'vitest';

const captured: { onOptions?: () => void }[] = [];
vi.mock('../../../src/interact/error-menu.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../../src/interact/error-menu.js')>();
  return {
    ...actual,
    attachErrorMenu: (deps: Parameters<typeof actual.attachErrorMenu>[0]) => {
      captured.push(deps);
      return actual.attachErrorMenu(deps);
    },
  };
});

const { mountStubSheet } = await import('../../test-utils/mount.js');

describe('error menu Options wiring', () => {
  afterEach(() => {
    captured.length = 0;
    document.body.replaceChildren();
  });

  it('opens Preferences on the Error Checking pane from the mounted menu', async () => {
    const sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    try {
      const deps = captured.at(-1);
      expect(deps?.onOptions).toBeTypeOf('function');
      deps?.onOptions?.();
      const pane = document.querySelector<HTMLElement>(
        '.fc-prefs:not([hidden]) .fc-prefs__pane:not([hidden])[data-pane="errorChecking"]',
      );
      expect(pane).not.toBeNull();
    } finally {
      sheet.dispose();
    }
  });
});
