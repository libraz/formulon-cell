import { describe, expect, it } from 'vitest';
import {
  hydrateValidationsFromEngine,
  syncValidationsToEngine,
} from '../../../src/engine/validation-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { validationAt } from '../../../src/store/range-validations.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const entry = (ranges: [number, number, number, number][], extra: Record<string, unknown>) => ({
  ranges: ranges.map(([r0, c0, r1, c1]) => ({ sheet: 0, r0, c0, r1, c1 })),
  type: 2,
  op: 0,
  errorStyle: 0,
  allowBlank: true,
  showInputMessage: true,
  showErrorMessage: true,
  formula1: '1',
  formula2: '',
  ...extra,
});

/** Save, reload, hydrate, flush again and return both engine rule lists. */
async function roundTrip(source: WorkbookHandle) {
  const restored = await WorkbookHandle.loadBytes(source.save());
  const store = createSpreadsheetStore();
  hydrateValidationsFromEngine(restored, store, 0);
  const before = restored.getValidationsForSheet(0);
  syncValidationsToEngine(restored, store, 0);
  return { restored, store, before, after: restored.getValidationsForSheet(0) };
}

const shape = (rules: ReturnType<WorkbookHandle['getValidationsForSheet']>) =>
  rules.map((r) => ({
    ranges: r.ranges.map((x) => [x.r0, x.c0, x.r1, x.c1]),
    type: r.type,
    op: r.op,
    f1: r.formula1,
    f2: r.formula2,
  }));

describe.skipIf(!canLoadWasm())('data validation round-trip', () => {
  it('keeps formula-bound limits', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(
        wb.addValidationEntry(0, entry([[4, 3, 6, 3]], { formula1: 'A1', formula2: '$B$1+1' })),
      ).toBe(true);
      const { restored, store, before, after } = await roundTrip(wb);
      try {
        expect(before).toHaveLength(1);
        const v = validationAt(store.getState().format, { sheet: 0, row: 5, col: 3 });
        expect(v).toMatchObject({
          kind: 'decimal',
          op: 'between',
          aFormula: 'A1',
          bFormula: '$B$1+1',
        });
        expect(shape(after)).toEqual(shape(before));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps list sources given as a name or formula', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.addValidationEntry(0, entry([[0, 0, 0, 0]], { type: 3, op: 0, formula1: 'MyList' }));
      wb.addValidationEntry(
        0,
        entry([[1, 0, 1, 0]], { type: 3, op: 0, formula1: 'INDIRECT($D$1)' }),
      );
      wb.addValidationEntry(0, entry([[2, 0, 2, 0]], { type: 3, op: 0, formula1: '$D$1:$D$5' }));
      const { restored, store, before, after } = await roundTrip(wb);
      try {
        const fmt = store.getState().format;
        expect(validationAt(fmt, { sheet: 0, row: 0, col: 0 })).toMatchObject({
          kind: 'list',
          source: { ref: 'MyList' },
        });
        expect(validationAt(fmt, { sheet: 0, row: 1, col: 0 })).toMatchObject({
          source: { ref: 'INDIRECT($D$1)' },
        });
        expect(before).toHaveLength(3);
        expect(shape(after)).toEqual(shape(before));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps whole-column rules as one range without per-cell entries', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.addValidationEntry(
        0,
        entry([[0, 0, 1048575, 0]], { type: 1, formula1: '0', formula2: '9' }),
      );
      const { restored, store, before, after } = await roundTrip(wb);
      try {
        const fmt = store.getState().format;
        expect(fmt.formats.size).toBe(0);
        expect(fmt.rangeValidations).toHaveLength(1);
        expect(validationAt(fmt, { sheet: 0, row: 500000, col: 0 })).toMatchObject({
          kind: 'whole',
        });
        expect(validationAt(fmt, { sheet: 0, row: 5, col: 1 })).toBeUndefined();
        expect(shape(after)).toEqual(shape(before));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('writes a multi-cell rule back as merged rectangles', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.addValidationEntry(
        0,
        entry(
          [
            [0, 0, 2, 1],
            [5, 3, 5, 3],
          ],
          { formula1: '0', formula2: '9' },
        ),
      );
      const { restored, before, after } = await roundTrip(wb);
      try {
        expect(shape(after)).toEqual(shape(before));
        expect(after).toHaveLength(1);
        expect(after[0]?.ranges).toHaveLength(2);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
