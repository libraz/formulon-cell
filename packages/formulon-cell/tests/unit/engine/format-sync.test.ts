import { describe, expect, it } from 'vitest';
import { hydrateCommentsAndHyperlinksFromEngine } from '../../../src/engine/format-sync.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

interface FakeCell {
  addr: { sheet: number; row: number; col: number };
  value: { kind: 'blank' };
  formula: string | null;
}

const makeFake = (opts: {
  comments?: boolean;
  hyperlinks?: boolean;
  cells?: FakeCell[];
  comments_data?: Record<string, { author: string; text: string }>;
  hyperlinks_data?: {
    row: number;
    col: number;
    target: string;
    display: string;
    tooltip: string;
  }[];
}): WorkbookHandle => {
  const cells = opts.cells ?? [];
  const commentsData = opts.comments_data ?? {};
  const hyperlinksData = opts.hyperlinks_data ?? [];
  return {
    capabilities: {
      comments: opts.comments ?? false,
      hyperlinks: opts.hyperlinks ?? false,
    },
    cells: function* (sheet: number) {
      for (const c of cells) if (c.addr.sheet === sheet) yield c;
    },
    getComment: (sheet: number, row: number, col: number) =>
      commentsData[`${sheet}:${row}:${col}`] ?? null,
    getHyperlinks: (_sheet: number) => hyperlinksData,
  } as unknown as WorkbookHandle;
};

describe('hydrateCommentsAndHyperlinksFromEngine', () => {
  it('no-op when neither capability is supported', () => {
    const store = createSpreadsheetStore();
    const wb = makeFake({});
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('seeds comment field on every populated cell with a non-empty comment', () => {
    const store = createSpreadsheetStore();
    const wb = makeFake({
      comments: true,
      cells: [
        { addr: { sheet: 0, row: 1, col: 2 }, value: { kind: 'blank' }, formula: null },
        { addr: { sheet: 0, row: 3, col: 4 }, value: { kind: 'blank' }, formula: null },
      ],
      comments_data: {
        '0:1:2': { author: 'a', text: 'hello' },
        // 3:4 has no comment
      },
    });
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    const formats = store.getState().format.formats;
    expect(formats.get(addrKey({ sheet: 0, row: 1, col: 2 }))?.comment).toBe('hello');
    expect(formats.get(addrKey({ sheet: 0, row: 3, col: 4 }))).toBeUndefined();
  });

  it('skips empty comment text', () => {
    const store = createSpreadsheetStore();
    const wb = makeFake({
      comments: true,
      cells: [{ addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'blank' }, formula: null }],
      comments_data: { '0:0:0': { author: 'a', text: '' } },
    });
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('seeds hyperlink field for every entry from getHyperlinks', () => {
    const store = createSpreadsheetStore();
    const wb = makeFake({
      hyperlinks: true,
      hyperlinks_data: [
        { row: 0, col: 0, target: 'https://example.com', display: '', tooltip: '' },
        { row: 5, col: 1, target: 'mailto:foo@bar', display: 'Foo', tooltip: '' },
      ],
    });
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    const formats = store.getState().format.formats;
    expect(formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.hyperlink).toBe(
      'https://example.com',
    );
    expect(formats.get(addrKey({ sheet: 0, row: 5, col: 1 }))?.hyperlink).toBe('mailto:foo@bar');
  });

  it('skips hyperlinks with empty target', () => {
    const store = createSpreadsheetStore();
    const wb = makeFake({
      hyperlinks: true,
      hyperlinks_data: [{ row: 0, col: 0, target: '', display: '', tooltip: '' }],
    });
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('preserves pre-existing format fields on the same cell', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set(addrKey({ sheet: 0, row: 0, col: 0 }), { bold: true });
      return { ...s, format: { formats } };
    });
    const wb = makeFake({
      hyperlinks: true,
      hyperlinks_data: [{ row: 0, col: 0, target: 'https://x', display: '', tooltip: '' }],
    });
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(true);
    expect(fmt?.hyperlink).toBe('https://x');
  });
});
