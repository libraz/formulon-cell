import { describe, expect, it } from 'vitest';
import { evaluateCfFromEngine } from '../../../src/engine/cf-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';

const KIND_COLOR_SCALE = 1;
const KIND_DATA_BAR = 2;
const KIND_DXF = 0;
const KIND_ICON = 3;

const c = (
  r: number,
  g: number,
  b: number,
  a = 255,
): { r: number; g: number; b: number; a: number } => ({ r, g, b, a });

const baseMatch = {
  kind: 0,
  priority: 0,
  dxfIdEngaged: false,
  dxfId: 0,
  color: c(0, 0, 0, 255),
  barLengthPct: 0,
  barAxisPositionPct: 0,
  barIsNegative: false,
  barFill: c(0, 0, 0, 255),
  barBorderEngaged: false,
  barBorder: c(0, 0, 0, 255),
  barGradient: false,
  iconSetName: 0,
  iconIndex: 0,
};

const fakeWb = (
  conditionalFormat: boolean,
  cells: { row: number; col: number; matches: Partial<typeof baseMatch>[] }[],
): WorkbookHandle => {
  return {
    capabilities: { conditionalFormat },
    evaluateCfRange: () =>
      cells.map((cell) => ({
        row: cell.row,
        col: cell.col,
        matches: cell.matches.map((m) => ({ ...baseMatch, ...m })),
      })),
  } as unknown as WorkbookHandle;
};

describe('evaluateCfFromEngine', () => {
  it('returns empty map when capability is off', () => {
    const wb = fakeWb(false, [
      { row: 0, col: 0, matches: [{ kind: KIND_COLOR_SCALE, color: c(255, 0, 0) }] },
    ]);
    expect(evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).size).toBe(0);
  });

  it('lifts ColorScale into overlay.fill as rgb()', () => {
    const wb = fakeWb(true, [
      { row: 1, col: 2, matches: [{ kind: KIND_COLOR_SCALE, color: c(10, 20, 30) }] },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    expect(out.get(addrKey({ sheet: 0, row: 1, col: 2 }))?.fill).toBe('rgb(10, 20, 30)');
  });

  it('emits rgba() when alpha is below 255', () => {
    const wb = fakeWb(true, [
      { row: 0, col: 0, matches: [{ kind: KIND_COLOR_SCALE, color: c(10, 20, 30, 128) }] },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill).toBe('rgba(10, 20, 30, 0.502)');
  });

  it('lifts DataBar to overlay.bar (0..1) and overlay.barColor', () => {
    const wb = fakeWb(true, [
      {
        row: 0,
        col: 0,
        matches: [{ kind: KIND_DATA_BAR, barLengthPct: 75, barFill: c(50, 100, 200) }],
      },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    const overlay = out.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(overlay?.bar).toBeCloseTo(0.75);
    expect(overlay?.barColor).toBe('rgb(50, 100, 200)');
  });

  it('clamps bar length to [0, 1]', () => {
    const wb = fakeWb(true, [
      { row: 0, col: 0, matches: [{ kind: KIND_DATA_BAR, barLengthPct: 150 }] },
      { row: 0, col: 1, matches: [{ kind: KIND_DATA_BAR, barLengthPct: -10 }] },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bar).toBe(1);
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 1 }))?.bar).toBe(0);
  });

  it('drops DifferentialFormat and IconSet matches (no overlay support yet)', () => {
    const wb = fakeWb(true, [
      { row: 0, col: 0, matches: [{ kind: KIND_DXF }, { kind: KIND_ICON }] },
    ]);
    expect(evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).size).toBe(0);
  });

  it('later matches in the same cell win for fill (priority order)', () => {
    const wb = fakeWb(true, [
      {
        row: 0,
        col: 0,
        matches: [
          { kind: KIND_COLOR_SCALE, color: c(255, 0, 0) }, // low priority
          { kind: KIND_COLOR_SCALE, color: c(0, 255, 0) }, // higher
        ],
      },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill).toBe('rgb(0, 255, 0)');
  });
});
