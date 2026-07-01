import { describe, expect, it, vi } from 'vitest';

import type { Extension } from '../../../src/extensions/types.js';
import { dedupeById, flattenExtensions, sortByPriority } from '../../../src/extensions/types.js';

const ext = (id: string, priority?: number): Extension => ({
  id,
  ...(priority !== undefined ? { priority } : {}),
  setup: () => ({ dispose: vi.fn() }),
});

describe('extension composition helpers', () => {
  it('flattens nested extension arrays without reordering leaves', () => {
    const a = ext('a');
    const b = ext('b');
    const c = ext('c');
    const d = ext('d');

    expect(flattenExtensions([a, [b, [c]], d]).map((entry) => entry.id)).toEqual([
      'a',
      'b',
      'c',
      'd',
    ]);
  });

  it('sorts by ascending priority while preserving input order for ties', () => {
    const late = ext('late', 80);
    const defaultA = ext('default-a');
    const early = ext('early', 10);
    const defaultB = ext('default-b');

    expect(sortByPriority([late, defaultA, early, defaultB]).map((entry) => entry.id)).toEqual([
      'early',
      'default-a',
      'default-b',
      'late',
    ]);
  });

  it('dedupes by id with last registration winning and preserves survivor order', () => {
    const firstFind = ext('findReplace', 10);
    const clipboard = ext('clipboard', 20);
    const secondFind = ext('findReplace', 90);
    const contextMenu = ext('contextMenu', 80);

    const result = dedupeById([firstFind, clipboard, secondFind, contextMenu]);

    expect(result).toEqual([clipboard, secondFind, contextMenu]);
  });

  it('supports the intended flatten then dedupe then sort pipeline', () => {
    const builtInFind = ext('findReplace', 50);
    const builtInClipboard = ext('clipboard', 50);
    const customFind = ext('findReplace', 10);
    const statusBar = ext('statusBar', 20);

    const result = sortByPriority(
      dedupeById(flattenExtensions([[builtInFind, builtInClipboard], [customFind], statusBar])),
    );

    expect(result.map((entry) => entry.id)).toEqual(['findReplace', 'statusBar', 'clipboard']);
    expect(result[0]).toBe(customFind);
  });
});
