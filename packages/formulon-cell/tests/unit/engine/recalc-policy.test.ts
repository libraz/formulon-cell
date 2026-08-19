// Recalc policy against the real WASM engine — the stub exposes neither
// `partialRecalc` nor `calcMode`, so these rules can only be pinned here.
import { describe, expect, it } from 'vitest';
import type { Addr } from '../../../src/engine/types.js';
import { type ChangeEvent, WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const A1: Addr = { sheet: 0, row: 0, col: 0 };
const NEAR: Addr = { sheet: 0, row: 1, col: 1 };
/** Far below anything a viewport would cover. */
const FAR: Addr = { sheet: 0, row: 500, col: 0 };

const newWb = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault({});
  expect(wb.isStub, 'these tests need the real engine').toBe(false);
  return wb;
};

const numberAt = (wb: WorkbookHandle, a: Addr): number | null => {
  const v = wb.getValue(a);
  return v.kind === 'number' ? v.value : null;
};

describe('recalc coverage', () => {
  it('recomputes dependents no matter how far they sit from the edit', async () => {
    const wb = await newWb();
    wb.setNumber(A1, 20);
    wb.setFormula(NEAR, '=A1*2');
    wb.setFormula(FAR, '=A1*3');
    expect(numberAt(wb, NEAR)).toBe(40);
    expect(numberAt(wb, FAR)).toBe(60);

    wb.setNumber(A1, 25);

    expect(numberAt(wb, NEAR)).toBe(50);
    // Regression: a viewport-scoped pass used to leave this one at 60.
    expect(numberAt(wb, FAR)).toBe(75);
    wb.dispose();
  });

  it('keeps partialRecalc available as an explicit, viewport-scoped call', async () => {
    const wb = await newWb();
    expect(wb.capabilities.partialRecalc).toBe(true);
    wb.setNumber(A1, 2);
    wb.setFormula(FAR, '=A1*3');
    expect(wb.partialRecalc(0, 0, 0, 20, 20)).not.toBeNull();
    expect(numberAt(wb, FAR)).toBe(6);
    wb.dispose();
  });
});

describe('manual calc mode', () => {
  it('holds edits until an explicit recalc', async () => {
    const wb = await newWb();
    expect(wb.capabilities.calcMode).toBe(true);
    wb.setNumber(A1, 20);
    wb.setFormula(NEAR, '=A1*2');
    expect(numberAt(wb, NEAR)).toBe(40);

    expect(wb.setCalcMode(1)).toBe(true);
    wb.setNumber(A1, 25);
    expect(numberAt(wb, NEAR)).toBe(40);

    // Calculate Now works from any mode.
    wb.recalc();
    expect(numberAt(wb, NEAR)).toBe(50);
    wb.dispose();
  });

  it('skips recalcAuto but not recalc', async () => {
    const wb = await newWb();
    wb.setNumber(A1, 1);
    wb.setFormula(NEAR, '=A1*2');
    wb.setCalcMode(1);

    wb.setNumber(A1, 4);
    wb.recalcAuto();
    expect(numberAt(wb, NEAR)).toBe(2);

    wb.recalc();
    expect(numberAt(wb, NEAR)).toBe(8);
    wb.dispose();
  });

  it('settles the sheet when the mode goes back to automatic', async () => {
    const wb = await newWb();
    wb.setNumber(A1, 3);
    wb.setFormula(NEAR, '=A1*2');
    wb.setCalcMode(1);
    wb.setNumber(A1, 7);
    expect(numberAt(wb, NEAR)).toBe(6);

    wb.setCalcMode(0);

    expect(numberAt(wb, NEAR)).toBe(14);
    wb.dispose();
  });

  it('resumes automatic recalc for later edits', async () => {
    const wb = await newWb();
    wb.setNumber(A1, 1);
    wb.setFormula(NEAR, '=A1*2');
    wb.setCalcMode(1);
    wb.setCalcMode(0);

    wb.setNumber(A1, 5);

    expect(numberAt(wb, NEAR)).toBe(10);
    wb.dispose();
  });
});

describe('recalc change event', () => {
  const recalcEvents = (wb: WorkbookHandle): { seen: ReadonlySet<string>[] } => {
    const seen: ReadonlySet<string>[] = [];
    wb.subscribe((e: ChangeEvent) => {
      if (e.kind === 'recalc') seen.push(e.dirty);
    });
    return { seen };
  };

  it('reports the cells written since the previous pass', async () => {
    const wb = await newWb();
    wb.setFormula(NEAR, '=A1*2');
    const { seen } = recalcEvents(wb);

    wb.setNumber(A1, 5);

    expect(seen).toHaveLength(1);
    expect([...(seen[0] as ReadonlySet<string>)]).toEqual(['0:0:0']);
    wb.dispose();
  });

  it('fires once for a batched write, listing every cell in it', async () => {
    const wb = await newWb();
    const { seen } = recalcEvents(wb);

    wb.withBatchedRecalc(() => {
      wb.setNumber(A1, 1);
      wb.setNumber(FAR, 2);
    });

    expect(seen).toHaveLength(1);
    expect([...(seen[0] as ReadonlySet<string>)].sort()).toEqual(['0:0:0', '0:500:0']);
    wb.dispose();
  });

  it('carries the held-back edits over to the pass that finally runs', async () => {
    const wb = await newWb();
    wb.setCalcMode(1);
    const { seen } = recalcEvents(wb);

    wb.setNumber(A1, 1);
    expect(seen).toHaveLength(0);

    wb.recalc();

    expect(seen).toHaveLength(1);
    expect([...(seen[0] as ReadonlySet<string>)]).toEqual(['0:0:0']);
    wb.dispose();
  });
});
