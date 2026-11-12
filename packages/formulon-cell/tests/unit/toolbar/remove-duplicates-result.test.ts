import { describe, expect, it } from 'vitest';
import { removeDuplicates } from '../../../src/commands/sort.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { de, en, ja } from '../../../src/i18n/strings.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { removeDuplicatesResultMessage } from '../../../src/toolbar/dialogs/remove-duplicates.js';

describe('Remove Duplicates result alert', () => {
  it('reports removed and remaining counts in the desktop wording', () => {
    expect(
      removeDuplicatesResultMessage({ ok: true, removed: 2, remaining: 3 }, en.ribbonMenu),
    ).toBe(
      '2 duplicate values found and removed; 3 unique values remain. Note that counts may include empty cells, spaces, etc.',
    );
    expect(
      removeDuplicatesResultMessage({ ok: true, removed: 1, remaining: 4 }, ja.ribbonMenu),
    ).toBe(
      '1 個の重複する値が見つかり、削除されました。4 個の一意の値が残ります。カウントには空のセルやスペースなどが含まれる場合があることに注意してください。',
    );
  });

  it('says none were found when nothing was removed', () => {
    expect(
      removeDuplicatesResultMessage({ ok: true, removed: 0, remaining: 5 }, de.ribbonMenu),
    ).toBe('Es wurden keine doppelten Werte gefunden.');
  });

  it('never reports a refusal as no duplicates found', () => {
    for (const reason of ['merged-cells', 'protected', 'too-large', 'engine-failed'] as const) {
      expect(
        removeDuplicatesResultMessage({ ok: false, reason }, en.ribbonMenu),
        reason,
      ).toBeNull();
    }
  });

  it('counts the unique data rows that remain, the header excluded', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      ['Name', 'a', 'b', 'a', 'c', 'b'].forEach((value, row) => {
        wb.setText({ sheet: 0, row, col: 0 }, value);
      });
      mutators.replaceCells(store, wb.cells(0));
      const result = removeDuplicates(
        store.getState(),
        store,
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 },
        { hasHeader: true },
      );
      expect(result).toEqual({ ok: true, removed: 2, remaining: 3 });
    } finally {
      wb.dispose();
    }
  });

  it('refuses a range with merged cells', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      for (const row of [0, 1, 2]) wb.setText({ sheet: 0, row, col: 0 }, 'x');
      mutators.replaceCells(store, wb.cells(0));
      store.setState((state) => {
        const byAnchor = new Map(state.merges.byAnchor);
        byAnchor.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
          sheet: 0,
          r0: 1,
          c0: 0,
          r1: 1,
          c1: 1,
        });
        return { ...state, merges: { ...state.merges, byAnchor } };
      });
      expect(
        removeDuplicates(store.getState(), store, wb, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }),
      ).toEqual({ ok: false, reason: 'merged-cells' });
    } finally {
      wb.dispose();
    }
  });
});
