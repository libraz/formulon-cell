import { describe, expect, it, vi } from 'vitest';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const A1: Addr = { sheet: 0, row: 0, col: 0 };
const A2: Addr = { sheet: 0, row: 1, col: 0 };
const B1: Addr = { sheet: 0, row: 0, col: 1 };
const B2: Addr = { sheet: 0, row: 1, col: 1 };

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const numberAt = (wb: WorkbookHandle, a: Addr): number | null => {
  const v = wb.getValue(a);
  return v.kind === 'number' ? v.value : null;
};

/** Count engine-level recalc passes. The handle deliberately keeps its engine
 *  private, so the spy is installed through a narrow cast rather than a public
 *  seam that production code could lean on. */
const spyEngineRecalc = (wb: WorkbookHandle): ReturnType<typeof vi.fn> => {
  const internals = wb as unknown as { wb: { recalc: () => { ok: boolean } } };
  const original = internals.wb.recalc.bind(internals.wb);
  const spy = vi.fn(original);
  internals.wb.recalc = spy;
  return spy;
};

describe('recalc on cell write', () => {
  it('recomputes dependent formulas after a number edit', async () => {
    const wb = await newWb();
    wb.setNumber(A1, 20);
    wb.setFormula(B1, '=A1*2');
    expect(numberAt(wb, B1)).toBe(40);

    wb.setNumber(A1, 25);

    expect(numberAt(wb, B1)).toBe(50);
    wb.dispose();
  });

  it('recomputes dependent formulas after a text edit replaces a number', async () => {
    const wb = await newWb();
    wb.setNumber(A1, 3);
    wb.setFormula(B1, '=A1*2');
    expect(numberAt(wb, B1)).toBe(6);

    wb.setText(A1, 'not a number');

    expect(numberAt(wb, B1)).not.toBe(6);
    wb.dispose();
  });

  it('recomputes dependent formulas after a precedent is blanked', async () => {
    const wb = await newWb();
    wb.setNumber(A1, 4);
    wb.setFormula(B1, '=A1*2');
    expect(numberAt(wb, B1)).toBe(8);

    wb.setBlank(A1);

    expect(numberAt(wb, B1)).toBe(0);
    wb.dispose();
  });
});

describe('withBatchedRecalc', () => {
  it('collapses a multi-cell write into one recalc and still lands the results', async () => {
    const wb = await newWb();
    wb.setFormula(B1, '=A1*2');
    wb.setFormula(B2, '=A2*2');
    const recalc = spyEngineRecalc(wb);

    wb.withBatchedRecalc(() => {
      wb.setNumber(A1, 5);
      wb.setNumber(A2, 6);
    });

    expect(recalc).toHaveBeenCalledTimes(1);
    expect(numberAt(wb, B1)).toBe(10);
    expect(numberAt(wb, B2)).toBe(12);
    wb.dispose();
  });

  it('flushes once at the outermost scope when batches nest', async () => {
    const wb = await newWb();
    wb.setFormula(B1, '=A1*2');
    const recalc = spyEngineRecalc(wb);

    wb.withBatchedRecalc(() => {
      wb.withBatchedRecalc(() => {
        wb.setNumber(A1, 7);
      });
      expect(recalc).not.toHaveBeenCalled();
      wb.setNumber(A2, 1);
    });

    expect(recalc).toHaveBeenCalledTimes(1);
    expect(numberAt(wb, B1)).toBe(14);
    wb.dispose();
  });

  it('still recalcs when the batched work throws', async () => {
    const wb = await newWb();
    wb.setFormula(B1, '=A1*2');
    const recalc = spyEngineRecalc(wb);

    expect(() =>
      wb.withBatchedRecalc(() => {
        wb.setNumber(A1, 9);
        throw new Error('boom');
      }),
    ).toThrow('boom');

    expect(recalc).toHaveBeenCalledTimes(1);
    expect(numberAt(wb, B1)).toBe(18);
    wb.dispose();
  });

  it('lets an explicit recalc inside the batch supersede the pending pass', async () => {
    const wb = await newWb();
    wb.setFormula(B1, '=A1*2');
    const recalc = spyEngineRecalc(wb);

    wb.withBatchedRecalc(() => {
      wb.setNumber(A1, 11);
      wb.recalc();
    });

    expect(recalc).toHaveBeenCalledTimes(1);
    expect(numberAt(wb, B1)).toBe(22);
    wb.dispose();
  });

  it('returns the callback result', async () => {
    const wb = await newWb();
    expect(wb.withBatchedRecalc(() => 'done')).toBe('done');
    wb.dispose();
  });
});
