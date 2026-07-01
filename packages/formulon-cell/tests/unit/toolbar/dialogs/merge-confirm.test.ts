import { afterEach, describe, expect, it } from 'vitest';
import type { Range } from '../../../../src/engine/types.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { confirmMergeLoseData } from '../../../../src/toolbar/dialogs/merge-confirm.js';

const RANGE: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };

const seedText = (store: ReturnType<typeof createSpreadsheetStore>, key: string): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(key, { value: { kind: 'text', value: 'x' }, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

const clickButton = (label: string): void => {
  const btn = [...document.body.querySelectorAll('button')].find((b) => b.textContent === label);
  btn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
};

afterEach(() => {
  document.body.innerHTML = '';
});

describe('confirmMergeLoseData', () => {
  it('resolves true without a dialog when no data would be lost', async () => {
    const store = createSpreadsheetStore();
    seedText(store, '0:0:0'); // only the anchor holds content
    const ok = await confirmMergeLoseData(defaultStrings, store.getState(), RANGE);
    expect(ok).toBe(true);
    expect(document.body.querySelector('button')).toBeNull();
  });

  it('resolves true when the user accepts the data-loss warning', async () => {
    const store = createSpreadsheetStore();
    seedText(store, '0:0:1'); // non-anchor content → warning
    const pending = confirmMergeLoseData(defaultStrings, store.getState(), RANGE);
    await Promise.resolve();
    clickButton(defaultStrings.ribbon.mergeLoseDataConfirm);
    expect(await pending).toBe(true);
  });

  it('resolves false when the user cancels', async () => {
    const store = createSpreadsheetStore();
    seedText(store, '0:0:1');
    const pending = confirmMergeLoseData(defaultStrings, store.getState(), RANGE);
    await Promise.resolve();
    clickButton(defaultStrings.ribbon.mergeLoseDataCancel);
    expect(await pending).toBe(false);
  });
});
