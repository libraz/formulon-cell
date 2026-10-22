// @vitest-environment node

import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  COLUMN_WIDTH_MAX_CHARS,
  type DimensionPrompt,
  ROW_HEIGHT_MAX_PT,
  runColumnWidthDialog,
  runRowHeightDialog,
} from '../../../src/commands/row-col-dimension.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const PX_PER_PT = 96 / 72;
const range = (r0: number, c0: number, r1: number, c1: number): Range => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});

const handles: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of handles.splice(0)) wb.dispose();
});

const setup = async (platform: 'default' | 'mac') => {
  const wb = await WorkbookHandle.createDefault({ locale: 'en', platform });
  handles.push(wb);
  const store = createSpreadsheetStore();
  // The host sets this from the platform: Mac draws 1 pt as 1 px.
  if (platform === 'mac') {
    store.setState((s) => ({ ...s, layout: { ...s.layout, screenScale: 0.75 } }));
  }
  hydrateLayoutFromEngine(wb, store, 0);
  return { wb, store };
};

const answering = (value: number | null) => vi.fn<DimensionPrompt>(async () => value);

describe('runRowHeightDialog', () => {
  it('shows points and snaps the typed height to whole 96-DPI pixels', async () => {
    const { wb, store } = await setup('default');
    const prompt = answering(15.3);
    await runRowHeightDialog({ store, history: null, wb, range: range(1, 0, 2, 0), prompt });
    expect(prompt).toHaveBeenCalledWith(15, ROW_HEIGHT_MAX_PT);
    expect(store.getState().layout.rowHeights.get(1)).toBe(20);
    expect(store.getState().layout.rowHeights.get(2)).toBe(20);
  });

  it('snaps to whole points on Mac, where a point is one screen pixel', async () => {
    const { wb, store } = await setup('mac');
    const prompt = answering(20.4);
    await runRowHeightDialog({ store, history: null, wb, range: range(0, 0, 0, 0), prompt });
    expect(prompt).toHaveBeenCalledWith(16, ROW_HEIGHT_MAX_PT);
    expect(store.getState().layout.rowHeights.get(0)).toBeCloseTo(20 * PX_PER_PT, 8);
  });

  it('leaves the field empty for rows of different heights', async () => {
    const { wb, store } = await setup('default');
    mutators.setRowHeight(store, 3, 30);
    const prompt = answering(null);
    await runRowHeightDialog({ store, history: null, wb, range: range(2, 0, 3, 0), prompt });
    expect(prompt).toHaveBeenCalledWith(undefined, ROW_HEIGHT_MAX_PT);
  });

  it('opens with the shared override when every row carries it', async () => {
    const { wb, store } = await setup('default');
    mutators.setRowHeight(store, 2, 30);
    mutators.setRowHeight(store, 3, 30);
    const prompt = answering(null);
    await runRowHeightDialog({ store, history: null, wb, range: range(2, 0, 3, 0), prompt });
    expect(prompt).toHaveBeenCalledWith(22.5, ROW_HEIGHT_MAX_PT);
  });

  it('hides the rows when the height is zero', async () => {
    const { wb, store } = await setup('default');
    await runRowHeightDialog({
      store,
      history: null,
      wb,
      range: range(4, 0, 5, 0),
      prompt: answering(0),
    });
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([4, 5]);
  });
});

describe('runColumnWidthDialog', () => {
  it('shows the Normal font characters and sets whole pixels on Windows', async () => {
    const { wb, store } = await setup('default');
    const prompt = answering(10);
    await runColumnWidthDialog({ store, history: null, wb, range: range(0, 1, 0, 1), prompt });
    expect(prompt).toHaveBeenCalledWith(8.43, COLUMN_WIDTH_MAX_CHARS);
    expect(store.getState().layout.colWidths.get(1)).toBe(75);
  });

  it('scales a sub-character width by digit plus padding', async () => {
    const { wb, store } = await setup('default');
    await runColumnWidthDialog({
      store,
      history: null,
      wb,
      range: range(0, 2, 0, 2),
      prompt: answering(0.5),
    });
    expect(store.getState().layout.colWidths.get(2)).toBe(6);
    expect(wb.columnDisplayChars(0, 6)).toBe(0.5);
  });

  it('works in points on Mac', async () => {
    const { wb, store } = await setup('mac');
    // Aptos Narrow 12 on Mac: a 6 pt digit, 5 pt padding, 65 pt default column.
    const prompt = answering(12.57);
    await runColumnWidthDialog({ store, history: null, wb, range: range(0, 0, 0, 0), prompt });
    expect(prompt).toHaveBeenCalledWith(10, COLUMN_WIDTH_MAX_CHARS);
    expect(store.getState().layout.colWidths.get(0)).toBeCloseTo(80 * PX_PER_PT, 8);
  });

  it('hides the columns when the width is zero', async () => {
    const { wb, store } = await setup('default');
    await runColumnWidthDialog({
      store,
      history: null,
      wb,
      range: range(0, 3, 0, 3),
      prompt: answering(0),
    });
    expect(store.getState().layout.hiddenCols.has(3)).toBe(true);
  });
});
