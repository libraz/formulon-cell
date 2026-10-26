import { describe, expect, it } from 'vitest';
import { coerceInputForCell } from '../../../src/commands/coerce-input.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { formatNumber } from '../../../src/format/number-format.js';
import { mountStubSheet } from '../../test-utils/mount.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe.skipIf(!canLoadWasm())('1904 workbook date interpretation', () => {
  it('formats and parses dates in the loaded workbook date system', async () => {
    const source = await WorkbookHandle.createDefault();
    let bytes: Uint8Array;
    try {
      expect(source.setCalcSetting('date1904', true)).toBe(true);
      source.setNumber({ sheet: 0, row: 0, col: 0 }, 43830);
      bytes = source.save();
    } finally {
      source.dispose();
    }
    const workbook = await WorkbookHandle.loadBytes(bytes);
    const sheet = await mountStubSheet({ workbook });
    try {
      const state = sheet.instance.store.getState();
      expect(state.ui.date1904).toBe(true);
      const cell = state.data.cells.get(addrKey({ sheet: 0, row: 0, col: 0 }));
      expect(cell?.value).toEqual({ kind: 'number', value: 43830 });
      const fmt = { kind: 'date', pattern: 'yyyy-mm-dd' } as const;
      expect(formatNumber(43830, fmt, 'en-US', state.ui.date1904 === true)).toBe('2024-01-01');
      expect(coerceInputForCell(state, { sheet: 0, row: 1, col: 0 }, '2024-01-01')).toMatchObject({
        kind: 'number',
        value: 43830,
      });
    } finally {
      sheet.dispose();
      workbook.dispose();
    }
  });
});
