import { afterEach, describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { prepareAxisFormatOverrideWithWorkbook } from '../../../src/engine/axis-format-sync.js';
import type { Addr, CellXf, FontRecord } from '../../../src/engine/types.js';
import type { WorkbookHandle as WorkbookHandleType } from '../../../src/engine/workbook-handle.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const rowRange = [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] as const;
const rowBand = [{ axis: 'row', first: 0, last: 0 }] as const;
const patch = { fontSize: 18 };

interface NativeCellWithXf {
  readonly addr: Addr;
}

interface TestWorkbook {
  readonly wb: WorkbookHandleType;
  readonly nativeCells: Addr[];
  readonly readNativeCells: ReturnType<typeof vi.fn>;
  readonly addedXfs: CellXf[];
}

const liveWorkbooks: WorkbookHandleType[] = [];

afterEach(() => {
  for (const wb of liveWorkbooks.splice(0)) wb.dispose();
});

const makeWorkbook = async (nativeCells: Addr[] = []): Promise<TestWorkbook> => {
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
  vi.spyOn(wb, 'getSheetFontAxisSnapshot').mockReturnValue({ columns: [], rows: [] });
  vi.spyOn(wb, 'getCellXfIndex').mockReturnValue(0);
  vi.spyOn(wb, 'getEffectiveStyle').mockReturnValue(null);
  vi.spyOn(wb, 'getCellXf').mockReturnValue(sourceXf);
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

  return { wb, nativeCells, readNativeCells, addedXfs };
};

const prepare = (
  wb: WorkbookHandleType,
  options: {
    exceptions?: readonly Addr[];
    resolvePatch?: Parameters<typeof prepareAxisFormatOverrideWithWorkbook>[7];
    budgetBase?: number;
  } = {},
) => {
  const store = createSpreadsheetStore();
  return prepareAxisFormatOverrideWithWorkbook(
    store.getState(),
    store,
    rowRange,
    rowBand,
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
