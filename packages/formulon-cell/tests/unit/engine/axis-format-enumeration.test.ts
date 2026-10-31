import { afterEach, describe, expect, it, type Mock, vi } from 'vitest';
import { addrKey, MAX_COL } from '../../../src/engine/address.js';
import {
  type AxisFormatBand,
  prepareAxisFormatOverrideWithWorkbook,
} from '../../../src/engine/axis-format-sync.js';
import type {
  Addr,
  CellXf,
  FontRecord,
  Range,
  SheetFontAxisSnapshot,
} from '../../../src/engine/types.js';
import type { WorkbookHandle as WorkbookHandleType } from '../../../src/engine/workbook-handle.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const rowRange = [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] as const;
const rowBand = [{ axis: 'row', first: 0, last: 0 }] as const;
const columnBand = [{ axis: 'column', first: 0, last: 0 }] as const;
const patch = { fontSize: 18 };

interface NativeCellWithXf {
  readonly addr: Addr;
}

type EffectiveStyle = NonNullable<ReturnType<WorkbookHandleType['getEffectiveStyle']>>;

interface TestWorkbook {
  readonly wb: WorkbookHandleType;
  readonly nativeCells: Addr[];
  readonly readNativeCells: Mock<(sheet: number) => Iterable<NativeCellWithXf>>;
  readonly effectiveResults: Array<EffectiveStyle | null>;
  readonly readEffectiveStyle: Mock<
    (sheet: number, row: number, col: number) => EffectiveStyle | null
  >;
  readonly readCellXfIndex: Mock<(sheet: number, row: number, col: number) => number | null>;
  readonly readXfIndices: number[];
  readonly addedXfs: CellXf[];
}

const liveWorkbooks: WorkbookHandleType[] = [];

afterEach(() => {
  for (const wb of liveWorkbooks.splice(0)) wb.dispose();
});

const defaultEffectiveStyle = { xfIndex: 7, source: 1 } as EffectiveStyle;

const makeWorkbook = async (
  nativeCells: Addr[] = [],
  options: { effectiveStyle?: EffectiveStyle | null; snapshot?: SheetFontAxisSnapshot } = {},
): Promise<TestWorkbook> => {
  const wb = await WorkbookHandle.createDefault({ preferStub: true });
  liveWorkbooks.push(wb);

  const sourceXf = {
    fontIndex: 0,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: 0,
    quotePrefix: true,
  } as CellXf;
  const addedXfs: CellXf[] = [];
  const effectiveResults: Array<EffectiveStyle | null> = [
    Object.hasOwn(options, 'effectiveStyle') ? (options.effectiveStyle ?? null) : null,
  ];
  const readEffectiveStyle = vi.fn((_sheet: number, _row: number, _col: number) => {
    return effectiveResults.length > 1
      ? (effectiveResults.shift() ?? null)
      : (effectiveResults[0] ?? null);
  });
  const readCellXfIndex = vi.fn((_sheet: number, _row: number, _col: number) => 0);
  const readXfIndices: number[] = [];
  const readXf = vi.fn((index: number) => {
    readXfIndices.push(index);
    return sourceXf;
  });
  vi.spyOn(wb, 'getSheetFontAxisSnapshot').mockReturnValue(
    options.snapshot ?? { columns: [], rows: [] },
  );
  vi.spyOn(wb, 'getCellXfIndex').mockImplementation(readCellXfIndex);
  vi.spyOn(wb, 'getEffectiveStyle').mockImplementation(readEffectiveStyle);
  vi.spyOn(wb, 'getCellXf').mockImplementation(readXf);
  vi.spyOn(wb, 'getFontRecord').mockReturnValue({ name: 'Aptos', size: 11 } as FontRecord);
  vi.spyOn(wb, 'getFillRecord').mockReturnValue(null);
  vi.spyOn(wb, 'getBorderRecord').mockReturnValue(null);
  vi.spyOn(wb, 'getNumFmtCode').mockReturnValue(null);
  vi.spyOn(wb, 'addFontRecord').mockReturnValue(1);
  vi.spyOn(wb, 'addFillRecord').mockReturnValue(1);
  vi.spyOn(wb, 'addBorderRecord').mockReturnValue(1);
  vi.spyOn(wb, 'addXfRecord').mockImplementation((xf) => {
    addedXfs.push(xf);
    return 100;
  });

  const readNativeCells = vi.fn(
    (_sheet: number): Iterable<NativeCellWithXf> => nativeCells.map((addr) => ({ addr })),
  );
  const extended = wb as WorkbookHandleType & {
    physicalCellsWithXf: (sheet: number) => Iterable<NativeCellWithXf>;
  };
  extended.physicalCellsWithXf = readNativeCells;

  return {
    wb,
    nativeCells,
    readNativeCells,
    effectiveResults,
    readEffectiveStyle,
    readCellXfIndex,
    readXfIndices,
    addedXfs,
  };
};

const prepare = (
  wb: WorkbookHandleType,
  options: {
    ranges?: readonly Range[];
    bands?: readonly AxisFormatBand[];
    exceptions?: readonly Addr[];
    resolvePatch?: Parameters<typeof prepareAxisFormatOverrideWithWorkbook>[7];
    budgetBase?: number;
  } = {},
) => {
  const store = createSpreadsheetStore();
  return prepareAxisFormatOverrideWithWorkbook(
    store.getState(),
    store,
    options.ranges ?? rowRange,
    options.bands ?? rowBand,
    patch,
    wb,
    options.exceptions,
    options.resolvePatch,
    options.budgetBase,
  );
};

describe('axis format native cell enumeration', () => {
  it.each([20, 1_000])(
    'reads %i selected native cells once and keeps the direct correction set',
    async (count) => {
      const addresses = Array.from({ length: count }, (_, index) => ({
        sheet: 0,
        row: 0,
        col: index,
      }));
      const fixture = await makeWorkbook(addresses);
      const result = prepare(fixture.wb);

      expect(fixture.readNativeCells).toHaveBeenCalledTimes(1);
      expect(result?.override.directXfs?.map((entry) => entry.key)).toEqual(addresses.map(addrKey));
      expect(result?.override.directXfs?.every((entry) => entry.afterXfIndex === 100)).toBe(true);
    },
  );

  it.each([20, 1_000])(
    'reads effective style once for each of %i valid native cells',
    async (count) => {
      const addresses = Array.from({ length: count }, (_, index) => ({
        sheet: 0,
        row: 0,
        col: index,
      }));
      const fixture = await makeWorkbook(addresses, { effectiveStyle: defaultEffectiveStyle });

      const result = prepare(fixture.wb);

      expect(fixture.readEffectiveStyle).toHaveBeenCalledTimes(count);
      expect(result?.override.directXfs?.map((entry) => entry.key)).toEqual(addresses.map(addrKey));
    },
  );

  it('performs no native cell enumeration when budget validation rejects early', async () => {
    const fixture = await makeWorkbook([{ sheet: 0, row: 0, col: 0 }]);
    const overBudget = prepare(fixture.wb, { budgetBase: 100_001 });
    const invalidException = prepare(fixture.wb, { exceptions: [{ sheet: 1, row: 0, col: 0 }] });

    expect(overBudget).toBeNull();
    expect(invalidException).toBeNull();
    expect(fixture.readNativeCells).not.toHaveBeenCalled();
  });

  it('rereads native cells after a resolver adds a cell during preparation', async () => {
    const first = { sheet: 0, row: 0, col: 0 } as const;
    const added = { sheet: 0, row: 0, col: 1 } as const;
    const fixture = await makeWorkbook([first]);
    let inserted = false;
    const resolvePatch = vi.fn(() => {
      if (!inserted) {
        fixture.nativeCells.push(added);
        inserted = true;
      }
      return patch;
    });

    const result = prepare(fixture.wb, { resolvePatch });

    expect(resolvePatch).toHaveBeenCalled();
    expect(fixture.readNativeCells).toHaveBeenCalledTimes(2);
    expect(result?.override.directXfs?.map((entry) => entry.key)).toEqual([
      addrKey(first),
      addrKey(added),
    ]);
  });

  it('keeps XF0 opaque metadata in the direct correction', async () => {
    const addr = { sheet: 0, row: 0, col: 0 } as const;
    const fixture = await makeWorkbook([addr]);
    const result = prepare(fixture.wb);

    expect(result?.override.directXfs).toEqual([
      expect.objectContaining({
        key: addrKey(addr),
        before: expect.objectContaining({ directXfIndex: 0 }),
      }),
    ]);
    expect(fixture.addedXfs.at(-1)?.quotePrefix).toBe(true);
  });

  it('retries a null first effective-style read for the patch base but keeps the original snapshot', async () => {
    const addr = { sheet: 0, row: 0, col: 0 } as const;
    const inherited = { xfIndex: 7, source: 1 } as EffectiveStyle;
    const fixture = await makeWorkbook([addr]);
    fixture.effectiveResults.splice(0, fixture.effectiveResults.length, null, inherited);

    const result = prepare(fixture.wb);

    expect(fixture.readEffectiveStyle).toHaveBeenCalledTimes(2);
    expect(fixture.readXfIndices).toContain(7);
    expect(result?.override.directXfs?.[0]).toMatchObject({
      key: addrKey(addr),
      before: { directXfIndex: 0, effectiveXfIndex: null, source: null },
    });
  });

  it('does not read effective style when the direct XF index is unavailable', async () => {
    const fixture = await makeWorkbook([{ sheet: 0, row: 0, col: 0 }]);
    fixture.readCellXfIndex.mockReturnValue(null);

    const result = prepare(fixture.wb);

    expect(result).toBeNull();
    expect(fixture.readEffectiveStyle).not.toHaveBeenCalled();
  });

  it('shares a styled-row effective read while retaining column precedence and skip rules', async () => {
    const rowSnapshot: SheetFontAxisSnapshot = {
      columns: [],
      rows: [{ row: 0, styleXf: 5 }],
    };
    const fixture = await makeWorkbook([], {
      snapshot: rowSnapshot,
      effectiveStyle: defaultEffectiveStyle,
    });

    const selectedRow = prepare(fixture.wb, { bands: columnBand });

    expect(selectedRow?.override.directXfs?.map((entry) => entry.key)).toEqual(['0:0:0']);
    expect(fixture.readEffectiveStyle).toHaveBeenCalledTimes(2);

    const skipped = await makeWorkbook([], {
      snapshot: rowSnapshot,
      effectiveStyle: { xfIndex: 8, source: 2 } as EffectiveStyle,
    });
    const skippedResult = prepare(skipped.wb, { bands: columnBand });

    expect(skippedResult?.override.directXfs).toBeUndefined();
    expect(skipped.readEffectiveStyle).toHaveBeenCalledTimes(1);
    expect(skipped.readCellXfIndex).not.toHaveBeenCalled();
  });

  it('shares effective reads across styled-column and derived-row intersection passes', async () => {
    const columnSnapshot: SheetFontAxisSnapshot = {
      columns: [
        { first: 0, last: 0, styleXf: 5 },
        { first: 1, last: MAX_COL, styleXf: 0 },
      ],
      rows: [],
    };
    const fixture = await makeWorkbook([], {
      snapshot: columnSnapshot,
      effectiveStyle: { xfIndex: 7, source: 2 } as EffectiveStyle,
    });

    const result = prepare(fixture.wb);

    expect(result?.override.directXfs?.map((entry) => entry.key)).toEqual(['0:0:0']);
    expect(fixture.readEffectiveStyle).toHaveBeenCalledTimes(3);
  });

  it('uses a fresh native cell snapshot for each separate preparation', async () => {
    const first = { sheet: 0, row: 0, col: 0 } as const;
    const added = { sheet: 0, row: 0, col: 1 } as const;
    const fixture = await makeWorkbook([first]);

    const before = prepare(fixture.wb);
    fixture.nativeCells.push(added);
    const after = prepare(fixture.wb);

    expect(fixture.readNativeCells).toHaveBeenCalledTimes(2);
    expect(before?.override.directXfs?.map((entry) => entry.key)).toEqual([addrKey(first)]);
    expect(after?.override.directXfs?.map((entry) => entry.key)).toEqual([
      addrKey(first),
      addrKey(added),
    ]);
  });
});
