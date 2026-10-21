import { describe, expect, it, vi } from 'vitest';
import {
  clearHyperlink,
  hyperlinkAt,
  hyperlinkCoverAt,
  isFollowableHyperlinkTarget,
  isSafeHyperlinkTarget,
  listEngineHyperlinks,
  listHyperlinks,
  setHyperlink,
} from '../../../src/commands/hyperlinks.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

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
    expect(addHyperlink).toHaveBeenCalledWith(0, 1, 1, 'https://x', '', '');
  });

  it('preserves display text and tooltip when only the URL is edited', () => {
    const store = createSpreadsheetStore();
    const clearHyperlinks = vi.fn(() => true);
    const addHyperlink = vi.fn(() => true);
    const wb = {
      capabilities: { hyperlinks: true },
      clearHyperlinks,
      addHyperlink,
    } as unknown as WorkbookHandle;
    const addr = { sheet: 0, row: 1, col: 1 };

    setHyperlink(store, addr, 'https://old', wb, {
      display: 'Display',
      tooltip: 'Tooltip',
    });
    setHyperlink(store, addr, 'https://new', wb);

    expect(listHyperlinks(store.getState(), 0)).toEqual([
      {
        addr,
        target: 'https://new',
        display: 'Display',
        tooltip: 'Tooltip',
      },
    ]);
    expect(addHyperlink).toHaveBeenLastCalledWith(0, 1, 1, 'https://new', 'Display', 'Tooltip');
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

describe('isSafeHyperlinkTarget', () => {
  it('allows http(s), mailto and tel regardless of case or surrounding whitespace', () => {
    for (const url of ['http://a.example', ' HTTPS://a.example ', 'mailto:a@b.c', 'tel:+81']) {
      expect(isSafeHyperlinkTarget(url)).toBe(true);
    }
  });

  it('rejects script-bearing and unknown schemes', () => {
    for (const url of [
      'javascript:alert(1)',
      'data:text/html,x',
      'file:///etc/passwd',
      '',
      'a.example',
    ]) {
      expect(isSafeHyperlinkTarget(url)).toBe(false);
    }
  });

  it('answers for every cell of a range link and clears it from any of them', () => {
    const store = createSpreadsheetStore();
    const anchor = { sheet: 0, row: 1, col: 1 };
    setHyperlink(store, anchor, 'https://span.example');
    mutators.setRangeFormat(
      store,
      { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      { hyperlinkSpan: { lastRow: 3, lastCol: 2 } },
    );

    expect(hyperlinkAt(store.getState(), { sheet: 0, row: 3, col: 2 })).toBe(
      'https://span.example',
    );
    expect(hyperlinkCoverAt(store.getState(), { sheet: 0, row: 2, col: 2 })?.addr).toEqual(anchor);
    expect(hyperlinkAt(store.getState(), { sheet: 0, row: 4, col: 2 })).toBeNull();
    expect(hyperlinkAt(store.getState(), { sheet: 1, row: 2, col: 2 })).toBeNull();

    clearHyperlink(store, { sheet: 0, row: 3, col: 2 });
    expect(hyperlinkAt(store.getState(), anchor)).toBeNull();
    expect(hyperlinkAt(store.getState(), { sheet: 0, row: 3, col: 2 })).toBeNull();
  });

  it('reads a location-only link as an in-document target', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { hyperlinkLocation: 'Sheet2!B2' });
    expect(hyperlinkAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBe('#Sheet2!B2');
  });

  it('treats #-prefixed targets as followable but not javascript:', () => {
    expect(isFollowableHyperlinkTarget('#Sheet2!A1')).toBe(true);
    expect(isFollowableHyperlinkTarget(' #Name ')).toBe(true);
    expect(isFollowableHyperlinkTarget('#')).toBe(false);
    expect(isFollowableHyperlinkTarget('javascript:alert(1)')).toBe(false);
  });
});
