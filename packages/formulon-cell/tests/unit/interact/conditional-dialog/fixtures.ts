import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { SpreadsheetStore } from '../../../../src/store/store.js';

export const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

export const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};
