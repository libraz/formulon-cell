import { describe, expect, it } from 'vitest';
import { syncHyperlinksToEngine } from '../../../src/engine/format-sync.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

interface AddCall {
  sheet: number;
  row: number;
  col: number;
  target: string;
  display: string;
  tooltip: string;
}

const makeFake = (
  opts: { hyperlinks?: boolean } = {},
): { wb: WorkbookHandle; clears: number[]; adds: AddCall[] } => {
  const clears: number[] = [];
  const adds: AddCall[] = [];
  const caps = { hyperlinks: opts.hyperlinks ?? true };
  const fake = {
    capabilities: caps,
    clearHyperlinks(sheet: number): boolean {
      clears.push(sheet);
      return caps.hyperlinks;
    },
    addHyperlink(
      sheet: number,
      row: number,
      col: number,
      target: string,
      display = '',
      tooltip = '',
    ): boolean {
      adds.push({ sheet, row, col, target, display, tooltip });
      return caps.hyperlinks;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, clears, adds };
};

describe('syncHyperlinksToEngine', () => {
  it('clears + writes one entry per cell with a hyperlink', () => {
    const { wb, clears, adds } = makeFake();
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [addrKey({ sheet: 0, row: 0, col: 0 }), { hyperlink: 'https://a.example' }],
          [addrKey({ sheet: 0, row: 1, col: 2 }), { hyperlink: 'https://b.example' }],
        ]),
      },
    }));
    syncHyperlinksToEngine(wb, store, 0);
    expect(clears).toEqual([0]);
    expect(adds).toHaveLength(2);
    expect(adds.map((a) => a.target).sort()).toEqual(['https://a.example', 'https://b.example']);
  });

  it('skips cells on other sheets', () => {
    const { wb, adds } = makeFake();
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[addrKey({ sheet: 1, row: 0, col: 0 }), { hyperlink: 'https://other' }]]),
      },
    }));
    syncHyperlinksToEngine(wb, store, 0);
    expect(adds).toHaveLength(0);
  });

  it('no-op when capability is off', () => {
    const { wb, clears, adds } = makeFake({ hyperlinks: false });
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[addrKey({ sheet: 0, row: 0, col: 0 }), { hyperlink: 'https://x' }]]),
      },
    }));
    syncHyperlinksToEngine(wb, store, 0);
    expect(clears).toEqual([]);
    expect(adds).toEqual([]);
  });
});
