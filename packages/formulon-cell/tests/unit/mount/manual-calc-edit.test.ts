import { afterEach, describe, expect, it } from 'vitest';
import { commitCellInput } from '../../../src/commands/cell-input-commit.js';
import { interactionControllerFor } from '../../../src/commands/interaction-controller.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const A2: Addr = { sheet: 0, row: 1, col: 0 };
const B2: Addr = { sheet: 0, row: 1, col: 1 };
const C2: Addr = { sheet: 0, row: 1, col: 2 };
const D2: Addr = { sheet: 0, row: 1, col: 3 };

describe('manual calculation holds a typed edit in a mounted sheet', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('leaves dependents stale until Calculate Now', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    sheet = await mountStubSheet({ workbook });
    const { instance } = sheet;
    workbook.setNumber(A2, 600);
    workbook.setNumber(B2, 3);
    workbook.setFormula(C2, '=A2*B2');
    workbook.setFormula(D2, '=C2*1.1');
    await Promise.resolve();
    const num = (a: Addr): number | null => {
      const v = workbook.getValue(a);
      return v.kind === 'number' ? v.value : null;
    };
    const shown = (a: Addr) =>
      instance.store.getState().data.cells.get(`${a.sheet}:${a.row}:${a.col}`)?.value;
    expect(num(C2)).toBe(1800);

    expect(workbook.setCalcMode(1)).toBe(true);
    expect(interactionControllerFor(instance.store)).toBeTruthy();
    const result = commitCellInput({
      store: instance.store,
      wb: workbook,
      addr: B2,
      raw: '5',
      origin: 'editor',
    });
    expect(result.status).toBe('applied');
    await Promise.resolve();

    expect(num(B2)).toBe(5);
    expect(num(C2)).toBe(1800);
    expect(num(D2)).toBeCloseTo(1980);
    expect(shown(C2)).toEqual({ kind: 'number', value: 1800 });

    instance.recalc();
    await Promise.resolve();
    expect(num(C2)).toBe(3000);
    expect(num(D2)).toBeCloseTo(3300);
    expect(shown(C2)).toEqual({ kind: 'number', value: 3000 });
  });
});
