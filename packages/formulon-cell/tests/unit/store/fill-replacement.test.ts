import { describe, expect, it } from 'vitest';
import { applySelectionFormatPatch, setFillColor } from '../../../src/commands/format.js';
import { patchXfIndexFromBase } from '../../../src/engine/cell-format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { formatWithPending } from '../../../src/store/pending-format.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const B1 = { sheet: 0, row: 0, col: 1 };

const GRADIENT: NonNullable<CellFormat['fillGradient']> = {
  type: 'linear',
  degree: 0,
  left: 0,
  right: 0,
  top: 0,
  bottom: 0,
  stops: [
    { position: 0, color: '#ff0000' },
    { position: 1, color: '#0000ff' },
  ],
};

const formatAt = (store: SpreadsheetStore, addr = A1): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey(addr));

const selectA1 = (store: SpreadsheetStore): void => {
  store.setState((s) => ({
    ...s,
    selection: { active: A1, anchor: A1, range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 } },
  }));
};

describe('a fill kind replaces the others', () => {
  it('setCellFormat: a solid colour drops a gradient, a gradient drops the solid and pattern', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { fillGradient: GRADIENT, bold: true });
    mutators.setCellFormat(store, A1, { fill: '#00ff00' });
    expect(formatAt(store)?.fillGradient).toBeUndefined();
    expect(formatAt(store)).toMatchObject({ fill: '#00ff00', bold: true });

    mutators.setCellFormat(store, A1, { fillPattern: 'gray50', fillPatternColor: '#000000' });
    mutators.setCellFormat(store, A1, { fillGradient: GRADIENT });
    expect(formatAt(store)?.fill).toBeUndefined();
    expect(formatAt(store)?.fillPattern).toBeUndefined();
    expect(formatAt(store)?.fillPatternColor).toBeUndefined();
    expect(formatAt(store)?.fillGradient).toEqual(GRADIENT);
  });

  it('setCellFormat: clearing the fill or setting a pattern drops a gradient', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { fillGradient: GRADIENT });
    mutators.setCellFormat(store, A1, { fill: undefined });
    expect(formatAt(store)?.fillGradient).toBeUndefined();

    mutators.setCellFormat(store, A1, { fillGradient: GRADIENT });
    mutators.setCellFormat(store, A1, { fillPattern: 'darkGrid' });
    expect(formatAt(store)?.fillGradient).toBeUndefined();
  });

  it('setRangeFormat follows the same rule', () => {
    const store = createSpreadsheetStore();
    mutators.setRangeFormat(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      { fillGradient: GRADIENT },
    );
    mutators.setRangeFormat(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, { fill: '#00ff00' });
    expect(formatAt(store, A1)?.fillGradient).toBeUndefined();
    expect(formatAt(store, B1)?.fillGradient).toBeUndefined();
  });

  it('selection commands and the pending format follow the same rule', () => {
    const store = createSpreadsheetStore();
    selectA1(store);
    mutators.setCellFormat(store, A1, { fillGradient: GRADIENT });
    applySelectionFormatPatch(
      store.getState(),
      store,
      { fill: '#123456' },
      { allowPending: false },
    );
    expect(formatAt(store)?.fillGradient).toBeUndefined();

    mutators.setCellFormat(store, A1, { fillGradient: GRADIENT });
    mutators.setPendingFormat(store, { addr: A1, format: { fill: '#654321' } });
    expect(formatWithPending(store.getState(), A1)?.fillGradient).toBeUndefined();

    mutators.setPendingFormat(store, null);
    mutators.setCellFormat(store, A1, { fillGradient: GRADIENT });
    setFillColor(store.getState(), store, null);
    expect(formatAt(store)?.fillGradient).toBeUndefined();
  });

  it('the engine writeback drops a loaded gradient when a solid colour is patched over it', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setText(A1, 'x');
      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, A1, { fillGradient: GRADIENT });
      const { syncCellFormatsToEngine } = await import('../../../src/engine/cell-format-sync.js');
      syncCellFormatsToEngine(wb, store, 0);
      const base = wb.getCellXfIndex(0, 0, 0) ?? -1;
      const patched = patchXfIndexFromBase(wb, base, { fill: '#00ff00' });
      const fill = wb.getFillRecord(wb.getCellXf(patched)?.fillIndex ?? -1);
      expect(fill?.gradient).toBeUndefined();
      expect(fill?.pattern).toBe(1);

      const regradient = patchXfIndexFromBase(wb, patched, { fillGradient: GRADIENT });
      expect(wb.getFillRecord(wb.getCellXf(regradient)?.fillIndex ?? -1)?.gradient?.degree).toBe(0);
    } finally {
      wb.dispose();
    }
  });
});
