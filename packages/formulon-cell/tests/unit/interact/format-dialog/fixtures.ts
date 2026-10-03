import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../../../../src/store/store.js';

export const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

export const setActive = (store: SpreadsheetStore, row: number, col: number): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row, col },
      anchor: { sheet: 0, row, col },
      range: { sheet: 0, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

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

export const flushRaf = (): Promise<void> =>
  new Promise<void>((resolve) => {
    requestAnimationFrame(() => resolve());
  });

export const mergeWorkbook = (): WorkbookHandle =>
  ({
    capabilities: { merges: true },
    engineClearMerges: () => true,
    engineAddMerge: () => true,
    setBlank: () => undefined,
    setText: () => undefined,
    getValue: () => ({ kind: 'blank' }),
  }) as unknown as WorkbookHandle;

export const seedText = (
  store: SpreadsheetStore,
  row: number,
  col: number,
  value: string,
): void => {
  mutators.setCell(store, { sheet: 0, row, col }, { kind: 'text', value });
};
