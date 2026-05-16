import { describe, expect, it, vi } from 'vitest';
import {
  clearHyperlink,
  hyperlinkAt,
  listEngineHyperlinks,
  listHyperlinks,
  setHyperlink,
} from '../../../src/commands/hyperlinks.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('hyperlink commands', () => {
  it('sets, reads, lists, and clears store-backed hyperlinks', () => {
    const store = createSpreadsheetStore();
    const a1 = { sheet: 0, row: 0, col: 0 };
    const c3 = { sheet: 0, row: 2, col: 2 };

    setHyperlink(store, c3, ' https://c.example ');
    setHyperlink(store, a1, 'https://a.example');

    expect(hyperlinkAt(store.getState(), a1)).toBe('https://a.example');
    expect(listHyperlinks(store.getState(), 0)).toEqual([
      { addr: a1, target: 'https://a.example' },
      { addr: c3, target: 'https://c.example' },
    ]);

    clearHyperlink(store, a1);
    expect(hyperlinkAt(store.getState(), a1)).toBeNull();
  });

  it('clearHyperlink is a no-op when the cell has no hyperlink', () => {
    const store = createSpreadsheetStore();
    const clearHyperlinks = vi.fn(() => true);
    const addHyperlink = vi.fn(() => true);
    const wb = {
      capabilities: { hyperlinks: true },
      clearHyperlinks,
      addHyperlink,
    } as unknown as WorkbookHandle;

    clearHyperlink(store, { sheet: 0, row: 0, col: 0 }, wb);

    expect(store.getState().format.formats.size).toBe(0);
    expect(clearHyperlinks).not.toHaveBeenCalled();
    expect(addHyperlink).not.toHaveBeenCalled();
  });

  it('lists engine-backed hyperlinks with addresses', () => {
    const wb = {
      getHyperlinks: () => [{ row: 1, col: 2, target: 'https://x', display: 'X', tooltip: 'tip' }],
    } as unknown as WorkbookHandle;

    expect(listEngineHyperlinks(wb, 3)).toEqual([
      {
        addr: { sheet: 3, row: 1, col: 2 },
        target: 'https://x',
        display: 'X',
        tooltip: 'tip',
      },
    ]);
  });

  it('flushes hyperlink changes to the engine when a workbook is provided', () => {
    const store = createSpreadsheetStore();
    const clearHyperlinks = vi.fn(() => true);
    const addHyperlink = vi.fn(() => true);
    const wb = {
      capabilities: { hyperlinks: true },
      clearHyperlinks,
      addHyperlink,
    } as unknown as WorkbookHandle;

    setHyperlink(store, { sheet: 0, row: 1, col: 1 }, 'https://x', wb);

    expect(clearHyperlinks).toHaveBeenCalledWith(0);
    expect(addHyperlink).toHaveBeenCalledWith(0, 1, 1, 'https://x');
  });

  it('skips setHyperlink on locked cells in protected sheets', () => {
    const store = createSpreadsheetStore();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const clearHyperlinks = vi.fn(() => true);
    const addHyperlink = vi.fn(() => true);
    const wb = {
      capabilities: { hyperlinks: true },
      clearHyperlinks,
      addHyperlink,
    } as unknown as WorkbookHandle;
    setProtectedSheet(store, 0, true);

    try {
      setHyperlink(store, { sheet: 0, row: 0, col: 0 }, 'https://blocked', wb);

      expect(hyperlinkAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
      expect(clearHyperlinks).not.toHaveBeenCalled();
      expect(addHyperlink).not.toHaveBeenCalled();
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it('skips clearHyperlink on locked cells in protected sheets', () => {
    const store = createSpreadsheetStore();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const clearHyperlinks = vi.fn(() => true);
    const wb = {
      capabilities: { hyperlinks: true },
      clearHyperlinks,
      addHyperlink: vi.fn(() => true),
    } as unknown as WorkbookHandle;
    setHyperlink(store, { sheet: 0, row: 0, col: 0 }, 'https://keep');
    setProtectedSheet(store, 0, true);

    try {
      clearHyperlink(store, { sheet: 0, row: 0, col: 0 }, wb);

      expect(hyperlinkAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBe('https://keep');
      expect(clearHyperlinks).not.toHaveBeenCalled();
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });
});
