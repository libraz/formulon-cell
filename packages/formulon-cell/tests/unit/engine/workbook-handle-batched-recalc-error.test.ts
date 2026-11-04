import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

describe('WorkbookHandle.withBatchedRecalc', () => {
  it('propagates the callback error when the settling recalc also throws', async () => {
    const wb = await WorkbookHandle.createDefault();
    open.push(wb);
    const internals = wb as unknown as { wb: { recalc: () => unknown } };
    const original = internals.wb.recalc.bind(internals.wb);
    const boom = new Error('fn failed');
    expect(() =>
      wb.withBatchedRecalc(() => {
        wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
        internals.wb.recalc = () => {
          throw new Error('recalc failed');
        };
        throw boom;
      }),
    ).toThrow(boom);
    internals.wb.recalc = original;
  });

  it('still recalcs after a normal batch', async () => {
    const wb = await WorkbookHandle.createDefault();
    open.push(wb);
    wb.withBatchedRecalc(() => {
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
      wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*3');
    });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 6 });
  });
});
