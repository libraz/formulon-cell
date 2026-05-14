import { describe, expect, it } from 'vitest';

import { addrKey } from '../../../src/engine/address.js';
import { evaluateCfFromEngine } from '../../../src/engine/cf-sync.js';
import { detectSpillRange, findSpillRanges } from '../../../src/engine/spill.js';
import type { CellValue } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

type CellMap = Map<string, { value: CellValue; formula: string | null }>;

const seedCells = (
  cells: { row: number; col: number; value: CellValue; formula?: string | null }[],
): CellMap => {
  const out: CellMap = new Map();
  for (const c of cells) {
    out.set(addrKey({ sheet: 0, row: c.row, col: c.col }), {
      value: c.value,
      formula: c.formula ?? null,
    });
  }
  return out;
};

/** Fake CF-aware workbook handle. The renderer-side flow:
 *  1. detect spill ranges from the cells map,
 *  2. ask the engine for CF matches over the bounding box,
 *  3. merge the two so every spilled cell gets its CF overlay.
 *
 *  We mock step 2 with a deterministic "value >= 5 → red" rule. */
const makeCfWb = (cellLookup: CellMap): WorkbookHandle => {
  return {
    capabilities: { conditionalFormat: true } as never,
    evaluateCfRange(
      _sheet: number,
      firstRow: number,
      firstCol: number,
      lastRow: number,
      lastCol: number,
    ): {
      row: number;
      col: number;
      matches: {
        kind: number;
        priority: number;
        dxfIdEngaged: boolean;
        dxfId: number;
        color: { r: number; g: number; b: number; a: number };
        barLengthPct: number;
        barAxisPositionPct: number;
        barIsNegative: boolean;
        barFill: { r: number; g: number; b: number; a: number };
        barBorderEngaged: boolean;
        barBorder: { r: number; g: number; b: number; a: number };
        barGradient: boolean;
        iconSetName: number;
        iconIndex: number;
      }[];
    }[] {
      const out: ReturnType<WorkbookHandle['evaluateCfRange']> = [];
      for (let r = firstRow; r <= lastRow; r += 1) {
        for (let c = firstCol; c <= lastCol; c += 1) {
          const cell = cellLookup.get(addrKey({ sheet: 0, row: r, col: c }));
          if (!cell || cell.value.kind !== 'number') continue;
          if (cell.value.value < 5) continue;
          out.push({
            row: r,
            col: c,
            matches: [
              {
                kind: 1, // KIND_COLOR_SCALE
                priority: 1,
                dxfIdEngaged: false,
                dxfId: 0,
                color: { r: 255, g: 0, b: 0, a: 255 },
                barLengthPct: 0,
                barAxisPositionPct: 0,
                barIsNegative: false,
                barFill: { r: 0, g: 0, b: 0, a: 0 },
                barBorderEngaged: false,
                barBorder: { r: 0, g: 0, b: 0, a: 0 },
                barGradient: false,
                iconSetName: 0,
                iconIndex: 0,
              },
            ],
          });
        }
      }
      return out;
    },
  } as unknown as WorkbookHandle;
};

describe('engine/spill × conditional-format intersection', () => {
  it('CF overlay reaches every cell of a freshly-detected spill range', () => {
    // A1 = `=SEQUENCE(4)` → spills downward into A2..A4 with values 1..4.
    // We then mutate A4 to 6 so it qualifies for the CF rule (>= 5).
    const cells = seedCells([
      { row: 0, col: 0, value: { kind: 'number', value: 1 }, formula: '=SEQUENCE(4)' },
      { row: 1, col: 0, value: { kind: 'number', value: 2 } },
      { row: 2, col: 0, value: { kind: 'number', value: 3 } },
      { row: 3, col: 0, value: { kind: 'number', value: 6 } },
    ]);
    const wb = makeCfWb(cells);

    const spills = findSpillRanges(cells, 0);
    expect(spills).toHaveLength(1);
    const r = spills[0];
    if (!r) throw new Error('expected one spill range');
    expect(r).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 });

    const overlay = evaluateCfFromEngine(wb, 0, r.r0, r.c0, r.r1, r.c1);
    // CF rule "value >= 5" hits only A4.
    expect(overlay.size).toBe(1);
    const a4Key = addrKey({ sheet: 0, row: 3, col: 0 });
    expect(overlay.get(a4Key)?.fill).toBe('rgb(255, 0, 0)');
  });

  it('spill expansion picks up newly-arrived rows so CF re-evaluates them', () => {
    // Start with a 2-row spill: anchor A1, value at A2. Detect → range r0..r1=0..1.
    const small = seedCells([
      { row: 0, col: 0, value: { kind: 'number', value: 7 }, formula: '=SEQUENCE(2)' },
      { row: 1, col: 0, value: { kind: 'number', value: 8 } },
    ]);
    const smallRange = detectSpillRange(small, 0, 0, 0);
    expect(smallRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });

    // Engine recalcs and adds A3..A4. Detector should follow the expanded shape.
    const big = seedCells([
      { row: 0, col: 0, value: { kind: 'number', value: 7 }, formula: '=SEQUENCE(4)' },
      { row: 1, col: 0, value: { kind: 'number', value: 8 } },
      { row: 2, col: 0, value: { kind: 'number', value: 9 } },
      { row: 3, col: 0, value: { kind: 'number', value: 10 } },
    ]);
    const bigRange = detectSpillRange(big, 0, 0, 0);
    expect(bigRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 });

    // CF rule (>= 5) over the expanded range catches every spilled cell.
    const wb = makeCfWb(big);
    const overlay = evaluateCfFromEngine(wb, 0, bigRange.r0, bigRange.c0, bigRange.r1, bigRange.c1);
    expect(overlay.size).toBe(4);
    for (let row = 0; row <= 3; row += 1) {
      expect(overlay.get(addrKey({ sheet: 0, row, col: 0 }))?.fill).toBe('rgb(255, 0, 0)');
    }
  });

  it('shrinking the spill drops CF overlays from cells that left the range', () => {
    // After a partial recalc the spill ends at A2 (was A4).
    const shrunk = seedCells([
      { row: 0, col: 0, value: { kind: 'number', value: 7 }, formula: '=SEQUENCE(2)' },
      { row: 1, col: 0, value: { kind: 'number', value: 8 } },
      // A3/A4 are now blank from the engine's perspective. Spill detector must
      // not include them.
      { row: 2, col: 0, value: { kind: 'blank' } },
      { row: 3, col: 0, value: { kind: 'blank' } },
    ]);
    const range = detectSpillRange(shrunk, 0, 0, 0);
    expect(range).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });

    const wb = makeCfWb(shrunk);
    const overlay = evaluateCfFromEngine(wb, 0, range.r0, range.c0, range.r1, range.c1);
    // Only A1 and A2 are in range; both >= 5 → both painted.
    expect(overlay.size).toBe(2);
    expect(overlay.has(addrKey({ sheet: 0, row: 2, col: 0 }))).toBe(false);
    expect(overlay.has(addrKey({ sheet: 0, row: 3, col: 0 }))).toBe(false);
  });

  it('does not query the engine when conditionalFormat capability is off', () => {
    const wb = {
      capabilities: { conditionalFormat: false } as never,
      evaluateCfRange: () => {
        throw new Error('should not be called when capability is off');
      },
    } as unknown as WorkbookHandle;

    const overlay = evaluateCfFromEngine(wb, 0, 0, 0, 1, 0);
    expect(overlay.size).toBe(0);
  });

  it('CF priority ordering: highest-priority match wins on a spilled cell', () => {
    const cells = seedCells([
      { row: 0, col: 0, value: { kind: 'number', value: 9 }, formula: '=SEQUENCE(1)' },
    ]);
    const wb: WorkbookHandle = {
      capabilities: { conditionalFormat: true } as never,
      evaluateCfRange(): ReturnType<WorkbookHandle['evaluateCfRange']> {
        return [
          {
            row: 0,
            col: 0,
            matches: [
              {
                kind: 1,
                priority: 1,
                dxfIdEngaged: false,
                dxfId: 0,
                color: { r: 100, g: 100, b: 100, a: 255 }, // first
                barLengthPct: 0,
                barAxisPositionPct: 0,
                barIsNegative: false,
                barFill: { r: 0, g: 0, b: 0, a: 0 },
                barBorderEngaged: false,
                barBorder: { r: 0, g: 0, b: 0, a: 0 },
                barGradient: false,
                iconSetName: 0,
                iconIndex: 0,
              },
              {
                kind: 1,
                priority: 2,
                dxfIdEngaged: false,
                dxfId: 0,
                color: { r: 200, g: 0, b: 0, a: 255 }, // second — overrides
                barLengthPct: 0,
                barAxisPositionPct: 0,
                barIsNegative: false,
                barFill: { r: 0, g: 0, b: 0, a: 0 },
                barBorderEngaged: false,
                barBorder: { r: 0, g: 0, b: 0, a: 0 },
                barGradient: false,
                iconSetName: 0,
                iconIndex: 0,
              },
            ],
          },
        ];
      },
    } as unknown as WorkbookHandle;

    const overlay = evaluateCfFromEngine(wb, 0, 0, 0, 0, 0);
    expect(overlay.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill).toBe('rgb(200, 0, 0)');
    // Avoid unused-cells lint.
    expect(cells.size).toBe(1);
  });
});
