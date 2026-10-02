import { describe, expect, it } from 'vitest';
import {
  hydrateCommentsAndHyperlinksFromEngine,
  syncHyperlinksToEngine,
} from '../../../src/engine/format-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
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
  comments_list?: { row: number; col: number; author: string; text: string }[];
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
  const commentsList = opts.comments_list;
  const hyperlinksData = opts.hyperlinks_data ?? [];
  const fake = {
    capabilities: {
      comments: opts.comments ?? false,
      commentsEnumerable: commentsList !== undefined,
      hyperlinks: opts.hyperlinks ?? false,
    },
    cells: function* (sheet: number) {
      for (const c of cells) if (c.addr.sheet === sheet) yield c;
    },
    getComment: (sheet: number, row: number, col: number) =>
      commentsData[`${sheet}:${row}:${col}`] ?? null,
    getHyperlinks: (_sheet: number) => hyperlinksData,
  };
  if (commentsList) {
    return {
      ...fake,
      getComments: (_sheet: number) => commentsList,
    } as unknown as WorkbookHandle;
  }
  return fake as unknown as WorkbookHandle;
};

const makeHyperlinkEngine = (opts: { clearResults?: boolean[]; addResults?: boolean[] } = {}) => {
  const clearCalls: number[] = [];
  const addCalls: { row: number; col: number; target: string; succeeded: boolean }[] = [];
  let clearIndex = 0;
  let addIndex = 0;
  const wb = {
    capabilities: { hyperlinks: true },
    clearHyperlinks(sheet: number): boolean {
      clearCalls.push(sheet);
      const result = opts.clearResults?.[clearIndex] ?? true;
      clearIndex += 1;
      return result;
    },
    addHyperlink(
      _sheet: number,
      row: number,
      col: number,
      target: string,
      _display: string,
      _tooltip: string,
    ): boolean {
      const succeeded = opts.addResults?.[addIndex] ?? true;
      addIndex += 1;
      addCalls.push({ row, col, target, succeeded });
      return succeeded;
    },
  };
  return { wb: wb as unknown as WorkbookHandle, clearCalls, addCalls };
};

const setHyperlinks = (
  store: ReturnType<typeof createSpreadsheetStore>,
  links: { row: number; col: number; target: string }[],
) => {
  store.setState((s) => ({
    ...s,
    format: {
      formats: new Map(
        links.map((link) => [
          addrKey({ sheet: 0, row: link.row, col: link.col }),
          { hyperlink: link.target },
        ]),
      ),
    },
  }));
};

describe('syncHyperlinksToEngine', () => {
  it('strict mode rejects clear failure while legacy mode remains best-effort', () => {
    const { wb } = makeHyperlinkEngine({ clearResults: [false, false] });
    const store = createSpreadsheetStore();
    setHyperlinks(store, [{ row: 0, col: 0, target: 'https://example.com' }]);

    expect(() => syncHyperlinksToEngine(wb, store, 0, { strict: true })).toThrow(
      /hyperlinks clearHyperlinks at sheet:0/,
    );
    expect(() => syncHyperlinksToEngine(wb, store, 0)).not.toThrow();
  });

  it('strict mode retries the complete hyperlink set after a partial add failure', () => {
    const { wb, clearCalls, addCalls } = makeHyperlinkEngine({
      addResults: [true, false, true, true],
    });
    const store = createSpreadsheetStore();
    setHyperlinks(store, [
      { row: 0, col: 0, target: 'https://one.example' },
      { row: 0, col: 1, target: 'https://two.example' },
    ]);

    expect(() => syncHyperlinksToEngine(wb, store, 0, { strict: true })).toThrow(
      /hyperlinks addHyperlink at 0:0:1/,
    );
    expect(() => syncHyperlinksToEngine(wb, store, 0, { strict: true })).not.toThrow();
    expect(clearCalls).toEqual([0, 0]);
    expect(addCalls).toHaveLength(4);
    expect(
      addCalls
        .slice(-2)
        .filter((call) => call.succeeded)
        .map((call) => call.target),
    ).toEqual(['https://one.example', 'https://two.example']);
  });

  it('strict mode is a no-op when hyperlink capability is unavailable', () => {
    const store = createSpreadsheetStore();
    const wb = {
      capabilities: { hyperlinks: false },
      clearHyperlinks: () => false,
      addHyperlink: () => false,
    } as unknown as WorkbookHandle;
    expect(() => syncHyperlinksToEngine(wb, store, 0, { strict: true })).not.toThrow();
  });
});

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
    expect(formats.get(addrKey({ sheet: 0, row: 5, col: 1 }))?.hyperlinkDisplay).toBe('Foo');
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
      return { ...s, format: { ...s.format, formats } };
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

  it('seeds comment author for round-trip writeback', () => {
    const store = createSpreadsheetStore();
    const wb = makeFake({
      comments: true,
      cells: [{ addr: { sheet: 0, row: 1, col: 2 }, value: { kind: 'blank' }, formula: null }],
      comments_data: {
        '0:1:2': { author: 'Alice', text: 'hello' },
      },
    });
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 2 }));
    expect(fmt?.comment).toBe('hello');
    expect(fmt?.commentAuthor).toBe('Alice');
  });

  it('hydrates comments on otherwise-empty cells when the engine can enumerate comments', () => {
    const store = createSpreadsheetStore();
    const wb = makeFake({
      comments: true,
      comments_list: [{ row: 9, col: 4, author: 'Alice', text: 'blank-cell note' }],
    });
    hydrateCommentsAndHyperlinksFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 9, col: 4 }));
    expect(fmt?.comment).toBe('blank-cell note');
    expect(fmt?.commentAuthor).toBe('Alice');
  });
});
