import { describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

interface RecalcCalls {
  recalc: number;
  partial: {
    sheet: number;
    firstRow: number;
    firstCol: number;
    lastRow: number;
    lastCol: number;
  }[];
  iterative: { enabled: boolean; maxIterations: number; maxChange: number }[];
  progress: (((it: number, mx: number, max: number) => boolean | undefined) | null)[];
}

/** Build a stand-in for `WorkbookHandle` that mirrors the real wrapper's
 *  capability gate. The shape exists so call sites can exercise the gate
 *  decision without spinning up the WASM module. */
const makeFake = (opts: {
  partialRecalc?: boolean;
  iterativeProgress?: boolean;
  partialResult?: { ok: boolean; recomputed: number; message?: string };
}): { wb: WorkbookHandle; calls: RecalcCalls } => {
  const calls: RecalcCalls = { recalc: 0, partial: [], iterative: [], progress: [] };
  const partialResult = opts.partialResult ?? { ok: true, recomputed: 0 };
  const caps = {
    partialRecalc: opts.partialRecalc ?? false,
    iterativeProgress: opts.iterativeProgress ?? false,
  };
  const recalc = (): void => {
    calls.recalc += 1;
  };
  const fake = {
    capabilities: caps,
    recalc,
    partialRecalc: (
      sheet: number,
      firstRow: number,
      firstCol: number,
      lastRow: number,
      lastCol: number,
    ): number | null => {
      if (!caps.partialRecalc) {
        recalc();
        return null;
      }
      calls.partial.push({ sheet, firstRow, firstCol, lastRow, lastCol });
      if (!partialResult.ok) throw new Error(partialResult.message ?? 'fail');
      return partialResult.recomputed;
    },
    setIterative: (enabled: boolean, maxIterations: number, maxChange: number): boolean => {
      if (!caps.iterativeProgress) return false;
      calls.iterative.push({ enabled, maxIterations, maxChange });
      return true;
    },
    setIterativeProgress: (
      cb:
        | ((iteration: number, maxResidual: number, maxIterations: number) => boolean | undefined)
        | null,
    ): boolean => {
      if (!caps.iterativeProgress) return false;
      calls.progress.push(cb);
      return true;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, calls };
};

describe('partialRecalc', () => {
  it('forwards the viewport rectangle when capability is on and returns recomputed count', () => {
    const { wb, calls } = makeFake({
      partialRecalc: true,
      partialResult: { ok: true, recomputed: 7 },
    });
    const n = wb.partialRecalc(0, 1, 2, 3, 4);
    expect(n).toBe(7);
    expect(calls.partial).toEqual([{ sheet: 0, firstRow: 1, firstCol: 2, lastRow: 3, lastCol: 4 }]);
    expect(calls.recalc).toBe(0);
  });

  it('falls back to full recalc and returns null when capability is off', () => {
    const { wb, calls } = makeFake({ partialRecalc: false });
    const n = wb.partialRecalc(0, 0, 0, 9, 9);
    expect(n).toBeNull();
    expect(calls.recalc).toBe(1);
    expect(calls.partial).toEqual([]);
  });
});

describe('setIterative / setIterativeProgress', () => {
  it('forwards arguments when capability is on', () => {
    const { wb, calls } = makeFake({ iterativeProgress: true });
    expect(wb.setIterative(true, 100, 0.001)).toBe(true);
    expect(calls.iterative).toEqual([{ enabled: true, maxIterations: 100, maxChange: 0.001 }]);
  });

  it('returns false (no-op) when iterative capability is off', () => {
    const { wb, calls } = makeFake({ iterativeProgress: false });
    expect(wb.setIterative(true, 100, 0.001)).toBe(false);
    expect(calls.iterative).toEqual([]);
  });

  it('installs and clears progress callback', () => {
    const { wb, calls } = makeFake({ iterativeProgress: true });
    const cb = (): void => {};
    expect(wb.setIterativeProgress(cb)).toBe(true);
    expect(wb.setIterativeProgress(null)).toBe(true);
    expect(calls.progress).toEqual([cb, null]);
  });

  it('progress callback is no-op when capability is off', () => {
    const { wb, calls } = makeFake({ iterativeProgress: false });
    expect(wb.setIterativeProgress(() => {})).toBe(false);
    expect(calls.progress).toEqual([]);
  });
});
