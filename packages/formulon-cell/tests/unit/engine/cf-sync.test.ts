import { describe, expect, it } from 'vitest';
import {
  evaluateCfFromEngine,
  hydrateConditionalRulesFromEngine,
} from '../../../src/engine/cf-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

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

  it('lifts DataBar to overlay.bar (0..1), axis, direction, and color', () => {
    const wb = fakeWb(true, [
      {
        row: 0,
        col: 0,
        matches: [
          {
            kind: KIND_DATA_BAR,
            barLengthPct: 75,
            barAxisPositionPct: 40,
            barIsNegative: true,
            barFill: c(50, 100, 200),
            barGradient: true,
          },
        ],
      },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    const overlay = out.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(overlay?.bar).toBeCloseTo(0.75);
    expect(overlay?.barAxis).toBeCloseTo(0.4);
    expect(overlay?.barDirection).toBe('left');
    expect(overlay?.barColor).toBe('rgb(50, 100, 200)');
    expect(overlay?.barGradient).toBe(true);
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

  it('drops DifferentialFormat matches because the dxf table is not exposed yet', () => {
    const wb = fakeWb(true, [{ row: 0, col: 0, matches: [{ kind: KIND_DXF }] }]);
    expect(evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).size).toBe(0);
  });

  it('lifts known IconSet ordinals into icon overlay metadata', () => {
    const wb = fakeWb(true, [
      { row: 0, col: 0, matches: [{ kind: KIND_ICON, iconSetName: 0, iconIndex: 2 }] },
      { row: 0, col: 1, matches: [{ kind: KIND_ICON, iconSetName: 1, iconIndex: 9 }] },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toMatchObject({
      iconKind: 'arrows3',
      iconSlot: 2,
    });
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 1 }))).toMatchObject({
      iconKind: 'arrows5',
      iconSlot: 4,
    });
  });

  it('ignores unknown IconSet ordinals', () => {
    const wb = fakeWb(true, [
      { row: 0, col: 0, matches: [{ kind: KIND_ICON, iconSetName: 99, iconIndex: 1 }] },
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

describe('hydrateConditionalRulesFromEngine', () => {
  const importedWb = (
    formats: ReturnType<WorkbookHandle['getConditionalFormats']>,
  ): WorkbookHandle =>
    ({
      capabilities: { conditionalFormatMutate: true },
      getConditionalFormats: () => formats,
    }) as unknown as WorkbookHandle;

  it('hydrates engine non-visual rules into store rules tagged by engine id', () => {
    const store = createSpreadsheetStore();
    hydrateConditionalRulesFromEngine(
      importedWb([
        {
          id: 'cf-cell',
          type: 1,
          priority: 1,
          stopIfTrue: true,
          sqref: [{ firstRow: 1, firstCol: 2, lastRow: 3, lastCol: 4 }],
          op: 5,
          formula1: '10',
        },
        {
          id: 'cf-formula',
          type: 0,
          priority: 2,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          formula1: 'A1>B1',
        },
        {
          id: 'cf-text',
          type: 7,
          priority: 3,
          stopIfTrue: false,
          sqref: [{ firstRow: 4, firstCol: 1, lastRow: 5, lastCol: 1 }],
          text: 'late',
        },
      ]),
      store,
      0,
    );

    expect(store.getState().conditional.rules).toEqual([
      {
        engineId: 'cf-cell',
        stopIfTrue: true,
        kind: 'cell-value',
        range: { sheet: 0, r0: 1, c0: 2, r1: 3, c1: 4 },
        op: '>',
        a: 10,
        apply: {},
      },
      {
        engineId: 'cf-formula',
        kind: 'formula',
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        formula: '=A1>B1',
        apply: {},
      },
      {
        engineId: 'cf-text',
        kind: 'text-contains',
        range: { sheet: 0, r0: 4, c0: 1, r1: 5, c1: 1 },
        text: 'late',
        apply: {},
      },
    ]);
  });

  it('replaces stale engine-hydrated rules for the sheet and preserves session rules', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      conditional: {
        rules: [
          {
            engineId: 'old',
            kind: 'cell-value',
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            op: '>',
            a: 1,
            apply: {},
          },
          {
            kind: 'cell-value',
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            op: '<',
            a: 9,
            apply: { bold: true },
          },
        ],
      },
    }));

    hydrateConditionalRulesFromEngine(
      importedWb([
        {
          id: 'new',
          type: 16,
          priority: 1,
          stopIfTrue: false,
          sqref: [{ firstRow: 1, firstCol: 1, lastRow: 3, lastCol: 1 }],
        },
      ]),
      store,
      0,
    );

    expect(store.getState().conditional.rules).toEqual([
      {
        kind: 'cell-value',
        range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        op: '<',
        a: 9,
        apply: { bold: true },
      },
      {
        engineId: 'new',
        kind: 'duplicates',
        range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 1 },
        apply: {},
      },
    ]);
  });

  it('hydrates top-bottom and average rules and expands multi-range sqref', () => {
    const store = createSpreadsheetStore();
    hydrateConditionalRulesFromEngine(
      importedWb([
        {
          id: 'top',
          type: 5,
          priority: 1,
          stopIfTrue: false,
          sqref: [
            { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 0 },
            { firstRow: 0, firstCol: 2, lastRow: 1, lastCol: 2 },
          ],
          rank: 3,
          percent: true,
          bottom: true,
        },
        {
          id: 'avg',
          type: 6,
          priority: 2,
          stopIfTrue: false,
          sqref: [{ firstRow: 5, firstCol: 0, lastRow: 9, lastCol: 0 }],
          aboveAverage: false,
          stdDev: 2,
        },
      ]),
      store,
      0,
    );

    expect(store.getState().conditional.rules).toMatchObject([
      { engineId: 'top', kind: 'top-bottom', mode: 'bottom', n: 3, percent: true },
      { engineId: 'top', kind: 'top-bottom', mode: 'bottom', n: 3, percent: true },
      { engineId: 'avg', kind: 'average', mode: 'below-std-dev', stdDev: 2 },
    ]);
  });
});
