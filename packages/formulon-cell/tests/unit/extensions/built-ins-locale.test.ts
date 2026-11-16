import { afterEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { findReplace, statusBar } from '../../../src/extensions/built-ins.js';
import type { BuiltInExtension, BuiltInExtensionHandle } from '../../../src/extensions/types.js';
import { de } from '../../../src/i18n/strings.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

let sheet: MountedStubSheet | undefined;

afterEach(() => {
  sheet?.dispose();
  sheet = undefined;
  document.body.replaceChildren();
});

// The overlay changes the title the dictionary lookup keys on, so only the
// supplied locale can identify the language.
const overlaidDe = { ...de, a11y: { ...de.a11y, spreadsheet: 'Custom grid title' } };

const setup = async (factory: () => unknown, statusbar = document.createElement('div')) => {
  sheet = await mountStubSheet({
    locale: 'de',
    features: { statusBar: false, findReplace: false },
  });
  const { instance, host } = sheet;
  const handle = (factory() as BuiltInExtension).setup({
    host,
    statusbar,
    store: instance.store,
    history: instance.history,
    i18n: { locale: 'de', strings: overlaidDe },
    workbook: instance.workbook,
    refreshCells: () => {},
  } as unknown as Parameters<BuiltInExtension['setup']>[0]) as BuiltInExtensionHandle;
  return { instance, host, handle, statusbar };
};

const seedNumbers = (store: MountedStubSheet['instance']['store'], values: number[]): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    values.forEach((value, row) => {
      cells.set(addrKey({ sheet: 0, row, col: 0 }), {
        value: { kind: 'number', value },
        formula: null,
      });
    });
    return {
      ...s,
      data: { ...s.data, cells },
      selection: {
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: values.length - 1, c1: 0 },
      },
    };
  });
};

describe('built-in extensions take the locale from the UI, not the dictionary text', () => {
  it('status bar aggregates use the UI locale decimal separator', async () => {
    const { instance, handle, statusbar } = await setup(statusBar);
    try {
      seedNumbers(instance.store, [1.25, 2.5]);
      (handle as unknown as { refresh?: () => void }).refresh?.();
      instance.store.setState((s) => ({ ...s }));
      expect(statusbar.querySelector('.fc-host__statusbar-aggs')?.textContent).toContain('3,75');
    } finally {
      handle.dispose();
    }
  });

  it('find matches the number as displayed in the UI locale', async () => {
    const { instance, host, handle } = await setup(findReplace);
    try {
      seedNumbers(instance.store, [1.5]);
      (handle as unknown as { open(): void }).open();
      const input = host.querySelector<HTMLInputElement>('.fc-find input[type="text"]');
      if (!input) throw new Error('find input missing');
      input.value = '1,5';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      expect(host.querySelector('.fc-find__pill')?.textContent).toBe('0 / 1');
    } finally {
      handle.dispose();
    }
  });
});
