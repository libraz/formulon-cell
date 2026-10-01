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
  barDirection: 0,
  iconSetName: 0,
  iconIndex: 0,
};

const fakeWb = (
  conditionalFormat: boolean,
  cells: { row: number; col: number; matches: Partial<typeof baseMatch>[] }[],
  dxfs: Record<number, NonNullable<ReturnType<WorkbookHandle['getDxf']>>> = {},
): WorkbookHandle => {
  return {
    capabilities: { conditionalFormat, conditionalFormatDxf: Object.keys(dxfs).length > 0 },
    evaluateCfRange: () =>
      cells.map((cell) => ({
        row: cell.row,
        col: cell.col,
        matches: cell.matches.map((m) => ({ ...baseMatch, ...m })),
      })),
    getDxf: (index: number) => dxfs[index] ?? null,
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
    expect(overlay?.bar).toBeCloseTo(0.3);
    expect(overlay?.barAxis).toBeCloseTo(0.4);
    expect(overlay?.barDirection).toBe('left');
    expect(overlay?.barColor).toBe('rgb(50, 100, 200)');
    expect(overlay?.barGradient).toBe(true);
  });

  it('uses the matching rule metadata for axis appearance while trusting match colors', () => {
    const wb = {
      ...fakeWb(true, [
        {
          row: 0,
          col: 0,
          matches: [
            {
              kind: KIND_DATA_BAR,
              priority: 7,
              barLengthPct: 50,
              barAxisPositionPct: 50,
              barIsNegative: true,
              barFill: c(192, 0, 0),
              barBorderEngaged: true,
              barBorder: c(127, 0, 0),
              barDirection: 1,
            },
          ],
        },
      ]),
      getConditionalFormats: () => [
        {
          id: 'bar-appearance',
          type: 3,
          priority: 7,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: c(0, 120, 212),
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            gradient: false,
            axisPosition: 1,
            border: c(31, 31, 31),
            axisColor: c(64, 64, 64),
            direction: 1,
          },
        },
      ],
      physicalCells: function* () {
        yield {
          addr: { sheet: 0, row: 0, col: 0 },
          value: { kind: 'number', value: -1 },
          formula: null,
        };
      },
    } as unknown as WorkbookHandle;

    const overlay = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).get(
      addrKey({ sheet: 0, row: 0, col: 0 }),
    );
    expect(overlay).toMatchObject({
      barColor: 'rgb(192, 0, 0)',
      barBorderColor: 'rgb(127, 0, 0)',
      barAxisColor: 'rgb(64, 64, 64)',
      barAxisVisible: true,
    });
  });

  it('draws a none-axis negative bar on the positive side with the same direction', () => {
    const wb = {
      ...fakeWb(true, [
        {
          row: 0,
          col: 0,
          matches: [
            {
              kind: KIND_DATA_BAR,
              priority: 99,
              barLengthPct: 60,
              barAxisPositionPct: 0,
              barIsNegative: true,
              barFill: c(192, 0, 0),
              barBorderEngaged: false,
              barDirection: 2,
            },
          ],
        },
      ]),
      getConditionalFormats: () => [
        {
          id: 'bar-none',
          type: 3,
          priority: 7,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: c(0, 120, 212),
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            axisPosition: 2,
            direction: 2,
          },
        },
      ],
    } as unknown as WorkbookHandle;

    const overlay = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).get(
      addrKey({ sheet: 0, row: 0, col: 0 }),
    );
    expect(overlay).toMatchObject({
      bar: 0.6,
      barAxis: 1,
      barDirection: 'left',
      barAxisVisible: false,
    });
  });

  it('hides an automatic metadata axis when the rule population has no negatives', () => {
    const wb = {
      ...fakeWb(true, [
        {
          row: 0,
          col: 0,
          matches: [
            {
              kind: KIND_DATA_BAR,
              priority: 3,
              barLengthPct: 50,
              barAxisPositionPct: 40,
              barIsNegative: false,
              barDirection: 1,
            },
          ],
        },
      ]),
      getConditionalFormats: () => [
        {
          id: 'bar-auto',
          type: 3,
          priority: 3,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          dataBar: {
            min: { type: 0, value: '0' },
            max: { type: 0, value: '100' },
            fill: c(0, 120, 212),
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            axisPosition: 0,
            direction: 1,
          },
        },
      ],
      physicalCells: function* () {
        yield {
          addr: { sheet: 0, row: 0, col: 0 },
          value: { kind: 'number', value: 10 },
          formula: null,
        };
      },
    } as unknown as WorkbookHandle;

    const overlay = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).get(
      addrKey({ sheet: 0, row: 0, col: 0 }),
    );
    expect(overlay?.barAxisVisible).toBe(false);
  });

  it('mirrors data-bar axis and signed side for explicit and context directions', () => {
    const wb = {
      ...fakeWb(true, [
        {
          row: 0,
          col: 0,
          matches: [
            {
              kind: KIND_DATA_BAR,
              barLengthPct: 20,
              barAxisPositionPct: 40,
              barIsNegative: false,
              barDirection: 0,
            },
          ],
        },
        {
          row: 0,
          col: 1,
          matches: [
            {
              kind: KIND_DATA_BAR,
              barLengthPct: 20,
              barAxisPositionPct: 40,
              barIsNegative: true,
              barDirection: 1,
            },
          ],
        },
        {
          row: 0,
          col: 2,
          matches: [
            {
              kind: KIND_DATA_BAR,
              barLengthPct: 20,
              barAxisPositionPct: 40,
              barIsNegative: true,
              barDirection: 2,
            },
          ],
        },
      ]),
      getSheetView: () => ({ rightToLeft: true }),
    } as unknown as WorkbookHandle;
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toMatchObject({
      barAxis: 0.6,
      barDirection: 'left',
    });
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 1 }))).toMatchObject({
      barAxis: 0.4,
      barDirection: 'left',
    });
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 2 }))).toMatchObject({
      barAxis: 0.6,
      barDirection: 'right',
    });
  });

  it('projects the signed engine bar length into the cell-width fraction', () => {
    const wb = {
      ...fakeWb(true, [
        {
          row: 0,
          col: 0,
          matches: [
            {
              kind: KIND_DATA_BAR,
              barLengthPct: 100,
              barAxisPositionPct: 33.333333,
              barIsNegative: true,
              barDirection: 1,
            },
          ],
        },
        {
          row: 0,
          col: 1,
          matches: [
            {
              kind: KIND_DATA_BAR,
              barLengthPct: 100,
              barAxisPositionPct: 33.333333,
              barIsNegative: false,
              barDirection: 1,
            },
          ],
        },
        {
          row: 0,
          col: 2,
          matches: [
            {
              kind: KIND_DATA_BAR,
              barLengthPct: 100,
              barAxisPositionPct: 33.333333,
              barIsNegative: true,
              barDirection: 2,
            },
          ],
        },
      ]),
      getSheetView: () => ({ rightToLeft: false }),
    } as unknown as WorkbookHandle;
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    const negativeLtr = out.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(negativeLtr?.bar).toBeCloseTo(1 / 3);
    expect(negativeLtr?.barAxis).toBeCloseTo(1 / 3);
    expect(negativeLtr?.barDirection).toBe('left');
    const positiveLtr = out.get(addrKey({ sheet: 0, row: 0, col: 1 }));
    expect(positiveLtr?.bar).toBeCloseTo(2 / 3);
    expect(positiveLtr?.barAxis).toBeCloseTo(1 / 3);
    expect(positiveLtr?.barDirection).toBe('right');
    const negativeRtl = out.get(addrKey({ sheet: 0, row: 0, col: 2 }));
    expect(negativeRtl?.bar).toBeCloseTo(1 / 3);
    expect(negativeRtl?.barAxis).toBeCloseTo(2 / 3);
    expect(negativeRtl?.barDirection).toBe('right');
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

  it('lifts DifferentialFormat matches through the dxf table', () => {
    const wb = fakeWb(
      true,
      [{ row: 0, col: 0, matches: [{ kind: KIND_DXF, dxfIdEngaged: true, dxfId: 2 }] }],
      {
        2: {
          fill: { pattern: 1, fgArgb: 0xff112233, bgArgb: 0 },
          font: {
            name: 'Calibri',
            size: 11,
            bold: true,
            italic: false,
            strike: false,
            underline: 0,
            vertAlign: 0,
            colorArgb: 0xff445566,
          },
        },
      },
    );
    expect(
      evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toEqual({
      fill: '#112233',
      bold: true,
      color: '#445566',
    });
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

  it('uses first-defined visual properties in priority order', () => {
    const wb = fakeWb(true, [
      {
        row: 0,
        col: 0,
        matches: [
          { kind: KIND_COLOR_SCALE, color: c(255, 0, 0) }, // highest priority
          { kind: KIND_COLOR_SCALE, color: c(0, 255, 0) }, // lower priority
        ],
      },
    ]);
    const out = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9);
    // The C API documents priority-ascending matches and first-defined
    // folding: lower-priority matches may fill omitted properties, but do
    // not overwrite a higher-priority property that is already defined.
    expect(out.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.fill).toBe('rgb(255, 0, 0)');
  });

  it('keeps the first overlapping data bar atomic, including an omitted border', () => {
    const wb = {
      ...fakeWb(true, [
        {
          row: 0,
          col: 0,
          matches: [
            {
              kind: KIND_DATA_BAR,
              priority: 1,
              barLengthPct: 25,
              barAxisPositionPct: 30,
              barIsNegative: false,
              barFill: c(255, 0, 0),
              barBorderEngaged: false,
              barDirection: 1,
            },
            {
              kind: KIND_DATA_BAR,
              priority: 2,
              barLengthPct: 80,
              barAxisPositionPct: 70,
              barIsNegative: false,
              barFill: c(0, 255, 0),
              barBorderEngaged: true,
              barBorder: c(0, 0, 255),
              barDirection: 2,
            },
          ],
        },
      ]),
      getConditionalFormats: () => [
        {
          id: 'bar-high',
          type: 3,
          priority: 1,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: c(255, 0, 0),
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            gradient: false,
            axisPosition: 1,
            direction: 1,
          },
        },
        {
          id: 'bar-low',
          type: 3,
          priority: 2,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: c(0, 255, 0),
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            gradient: true,
            axisPosition: 2,
            direction: 2,
            border: c(0, 0, 255),
          },
        },
      ],
    } as unknown as WorkbookHandle;

    const overlay = evaluateCfFromEngine(wb, 0, 0, 0, 9, 9).get(
      addrKey({ sheet: 0, row: 0, col: 0 }),
    );
    expect(overlay).toMatchObject({
      bar: 0.175,
      barAxis: 0.3,
      barDirection: 'right',
      barColor: 'rgb(255, 0, 0)',
      barAxisVisible: true,
      barGradient: false,
    });
    expect(overlay?.barBorderColor).toBeUndefined();
  });
});

describe('hydrateConditionalRulesFromEngine', () => {
  const importedWb = (
    formats: ReturnType<WorkbookHandle['getConditionalFormats']>,
    dxfs: Record<number, NonNullable<ReturnType<WorkbookHandle['getDxf']>>> = {},
  ): WorkbookHandle =>
    ({
      capabilities: {
        conditionalFormatMutate: true,
        conditionalFormatDxf: Object.keys(dxfs).length > 0,
      },
      getConditionalFormats: () => formats,
      getDxf: (index: number) => dxfs[index] ?? null,
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
        {
          id: 'cf-text-not',
          type: 8,
          priority: 4,
          stopIfTrue: false,
          sqref: [{ firstRow: 6, firstCol: 1, lastRow: 6, lastCol: 1 }],
          text: 'done',
        },
        {
          id: 'cf-text-begins',
          type: 9,
          priority: 5,
          stopIfTrue: false,
          sqref: [{ firstRow: 7, firstCol: 1, lastRow: 7, lastCol: 1 }],
          text: 'pre',
        },
        {
          id: 'cf-text-ends',
          type: 10,
          priority: 6,
          stopIfTrue: false,
          sqref: [{ firstRow: 8, firstCol: 1, lastRow: 8, lastCol: 1 }],
          text: 'post',
        },
        {
          id: 'cf-scale',
          type: 2,
          priority: 7,
          stopIfTrue: false,
          sqref: [{ firstRow: 9, firstCol: 1, lastRow: 9, lastCol: 3 }],
          colorScale: {
            thresholds: [{ type: 3 }, { type: 2, value: '50' }, { type: 4 }],
            colors: [
              { r: 255, g: 0, b: 0, a: 255 },
              { r: 255, g: 255, b: 255, a: 255 },
              { r: 0, g: 128, b: 0, a: 255 },
            ],
          },
        },
        {
          id: 'cf-bar',
          type: 3,
          priority: 8,
          stopIfTrue: false,
          sqref: [{ firstRow: 10, firstCol: 1, lastRow: 10, lastCol: 3 }],
          dataBar: {
            min: { type: 3, gte: true },
            max: { type: 4, gte: true },
            fill: { r: 0, g: 120, b: 212, a: 255 },
            showValue: false,
            minLengthPct: 0,
            maxLengthPct: 100,
            gradient: false,
            direction: 2,
          },
        },
        {
          id: 'cf-icons',
          type: 4,
          priority: 9,
          stopIfTrue: false,
          sqref: [{ firstRow: 11, firstCol: 1, lastRow: 11, lastCol: 3 }],
          iconSet: {
            name: 3,
            thresholds: [
              { type: 1, value: '33', gte: false },
              { type: 1, value: '67', gte: false },
            ],
            reverse: true,
            showValue: false,
            percent: true,
            floor: { type: 1, value: '10', gte: false },
          },
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
      {
        engineId: 'cf-text-not',
        kind: 'text-contains',
        range: { sheet: 0, r0: 6, c0: 1, r1: 6, c1: 1 },
        text: 'done',
        mode: 'not-contains',
        apply: {},
      },
      {
        engineId: 'cf-text-begins',
        kind: 'text-contains',
        range: { sheet: 0, r0: 7, c0: 1, r1: 7, c1: 1 },
        text: 'pre',
        mode: 'begins-with',
        apply: {},
      },
      {
        engineId: 'cf-text-ends',
        kind: 'text-contains',
        range: { sheet: 0, r0: 8, c0: 1, r1: 8, c1: 1 },
        text: 'post',
        mode: 'ends-with',
        apply: {},
      },
      {
        engineId: 'cf-scale',
        kind: 'color-scale',
        range: { sheet: 0, r0: 9, c0: 1, r1: 9, c1: 3 },
        stops: ['rgb(255, 0, 0)', 'rgb(255, 255, 255)', 'rgb(0, 128, 0)'],
        thresholds: [{ kind: 'min' }, { kind: 'percentile', value: 50 }, { kind: 'max' }],
      },
      {
        engineId: 'cf-bar',
        kind: 'data-bar',
        range: { sheet: 0, r0: 10, c0: 1, r1: 10, c1: 3 },
        color: 'rgb(0, 120, 212)',
        showValue: false,
        gradient: false,
        direction: 'right-to-left',
      },
      {
        engineId: 'cf-icons',
        kind: 'icon-set',
        range: { sheet: 0, r0: 11, c0: 1, r1: 11, c1: 3 },
        icons: 'traffic3',
        showValue: false,
        reverseOrder: true,
        thresholds: [
          { kind: 'percent', value: 33, gte: false },
          { kind: 'percent', value: 67, gte: false },
        ],
        floor: { kind: 'percent', value: 10, gte: false },
      },
    ]);
  });

  it('hydrates custom data-bar endpoints and direction metadata', () => {
    const store = createSpreadsheetStore();
    hydrateConditionalRulesFromEngine(
      importedWb([
        {
          id: 'cf-bar-custom',
          type: 3,
          priority: 1,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 2 }],
          dataBar: {
            min: { type: 0, value: '20', gte: false },
            max: { type: 2, value: '80', gte: true },
            fill: { r: 0, g: 120, b: 212, a: 255 },
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            direction: 2,
          },
        },
      ]),
      store,
      0,
    );

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'data-bar',
      min: { kind: 'number', value: 20, gte: false },
      max: { kind: 'percentile', value: 80, gte: true },
      direction: 'right-to-left',
    });
  });

  it('hydrates data-bar axis and sign-specific appearance metadata', () => {
    const store = createSpreadsheetStore();
    hydrateConditionalRulesFromEngine(
      importedWb([
        {
          id: 'cf-bar-appearance',
          type: 3,
          priority: 1,
          stopIfTrue: false,
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 2 }],
          dataBar: {
            min: { type: 3, gte: true },
            max: { type: 4, gte: true },
            fill: c(0, 120, 212),
            showValue: true,
            minLengthPct: 0,
            maxLengthPct: 100,
            gradient: true,
            axisPosition: 1,
            negativeFill: c(192, 0, 0),
            border: c(31, 31, 31),
            negativeBorder: c(127, 0, 0),
            axisColor: c(64, 64, 64),
            direction: 1,
          },
        },
      ]),
      store,
      0,
    );

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'data-bar',
      axisPosition: 'middle',
      negativeColor: 'rgb(192, 0, 0)',
      borderColor: 'rgb(31, 31, 31)',
      negativeBorderColor: 'rgb(127, 0, 0)',
      axisColor: 'rgb(64, 64, 64)',
    });
  });

  it('treats an empty engine id as imported and does not duplicate on rehydrate', () => {
    const store = createSpreadsheetStore();
    const formats = [
      {
        id: '',
        type: 1,
        priority: 1,
        stopIfTrue: false,
        sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        op: 5,
        formula1: '10',
      },
    ] satisfies ReturnType<WorkbookHandle['getConditionalFormats']>;

    hydrateConditionalRulesFromEngine(importedWb(formats), store, 0);
    hydrateConditionalRulesFromEngine(importedWb(formats), store, 0);

    expect(store.getState().conditional.rules).toHaveLength(1);
    expect(store.getState().conditional.rules[0]).toMatchObject({
      engineId: '',
      kind: 'cell-value',
    });
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

  it('hydrates dxf formatting into non-visual rule apply fields', () => {
    const store = createSpreadsheetStore();
    hydrateConditionalRulesFromEngine(
      importedWb(
        [
          {
            id: 'cf-dxf',
            type: 1,
            priority: 1,
            stopIfTrue: false,
            sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
            op: 5,
            formula1: '10',
            dxfId: 4,
          },
        ],
        {
          4: {
            fill: { pattern: 1, fgArgb: 0xffe2f0d9, bgArgb: 0 },
            font: {
              name: 'Calibri',
              size: 11,
              bold: true,
              italic: false,
              strike: false,
              underline: 0,
              vertAlign: 0,
              colorArgb: 0xff006100,
            },
          },
        },
      ),
      store,
      0,
    );

    expect(store.getState().conditional.rules[0]).toMatchObject({
      engineId: 'cf-dxf',
      kind: 'cell-value',
      apply: {
        fill: '#e2f0d9',
        bold: true,
        color: '#006100',
      },
    });
  });
});
