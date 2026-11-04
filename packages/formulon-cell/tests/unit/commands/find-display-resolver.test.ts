import { describe, expect, it } from 'vitest';
import { findAll, findDisplayResolver } from '../../../src/commands/find.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('findDisplayResolver', () => {
  it('scans the note entries once for any number of lookups', () => {
    const store = createSpreadsheetStore();
    let scans = 0;
    class CountingMap<K, V> extends Map<K, V> {
      override [Symbol.iterator](): MapIterator<[K, V]> {
        scans += 1;
        return super[Symbol.iterator]();
      }
    }
    const formats = new CountingMap<string, { comment: string }>();
    for (let row = 0; row < 20; row += 1) {
      formats.set(addrKey({ sheet: 0, row, col: 0 }), { comment: `note ${row}` });
    }
    store.setState((s) => ({ ...s, format: { ...s.format, formats } }));
    const state = store.getState();
    const opts = { query: 'note', lookIn: 'notes' as const };
    const matches = findAll(state, opts);
    expect(matches).toHaveLength(20);

    scans = 0;
    const display = findDisplayResolver(state, opts);
    const texts = matches.map((m) => display(m.addr));
    expect(texts[3]).toBe('note 3');
    expect(scans).toBe(1);
  });
});
