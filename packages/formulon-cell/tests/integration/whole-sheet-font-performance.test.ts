import { afterEach, describe, expect, it, vi } from 'vitest';
import { setFont } from '../../src/commands/format.js';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import type { FontRecord, SheetFontAxisSnapshot } from '../../src/engine/types.js';
import { type CellPatch, WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createDefaultRibbonHelpers } from '../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../src/mount.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/mount.js';

interface MountedFontToolbar {
  readonly workbook: WorkbookHandle;
  readonly sheet: MountedStubSheet;
  readonly toolbar: ReturnType<typeof Spreadsheet.mountToolbar>;
  readonly host: HTMLElement;
  dispose(): void;
}

const mountFontToolbar = async (populatedCells: number): Promise<MountedFontToolbar> => {
  const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
  if (workbook.isStub) throw new Error('native workbook required');
  const patches: CellPatch[] = Array.from({ length: populatedCells }, (_, index) => ({
    addr: { sheet: 0, row: Math.floor(index / 20), col: index % 20 },
    value: { kind: 'text', value: `cell-${index}` },
    formula: null,
  }));
  workbook.applyCellPatchAtomic(patches);
  const sheet = await mountStubSheet({ workbook, locale: 'ja' });
  const host = document.createElement('div');
  document.body.append(host);
  const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
    helpers: createDefaultRibbonHelpers(sheet.instance, { lang: 'ja' }),
  });
  return {
    workbook,
    sheet,
    toolbar,
    host,
    dispose: () => {
      toolbar.dispose();
      sheet.dispose();
      workbook.dispose();
      host.remove();
    },
  };
};

const fontAt = (workbook: WorkbookHandle, row: number, col: number): FontRecord => {
  const style = workbook.getEffectiveStyle(0, row, col);
  const xf = style && workbook.getCellXf(style.xfIndex);
  const font = xf && workbook.getFontRecord(xf.fontIndex);
  if (!font) throw new Error(`missing effective font at 0:${row}:${col}`);
  return font;
};

describe('whole-sheet font synchronization performance', () => {
  const mounted: MountedFontToolbar[] = [];
  afterEach(() => {
    for (const entry of mounted.splice(0)) entry.dispose();
  });

  it.each([20, 1_000])(
    'captures the first font baseline once and writes each of %i populated cell XFs once',
    async (populatedCells) => {
      const fixture = await mountFontToolbar(populatedCells);
      mounted.push(fixture);
      const { workbook, sheet, toolbar } = fixture;
      mutators.selectAll(sheet.instance.store);
      const snapshot = vi.spyOn(workbook, 'getSheetFontAxisSnapshot');
      const setCellXfIndex = vi.spyOn(workbook, 'setCellXfIndex');

      expect(toolbar.applyCommand('fontGrow')).toBe(true);

      expect(snapshot).toHaveBeenCalledTimes(1);
      expect(setCellXfIndex).toHaveBeenCalledTimes(populatedCells);
      expect(
        sheet.instance.store.getState().format.sheetFontOverrides?.get(0)?.baseline,
      ).toBeDefined();
      expect(fontAt(workbook, MAX_ROW, MAX_COL).size).toBe(12);

      snapshot.mockRestore();
      setCellXfIndex.mockRestore();
      expect(sheet.instance.undo()).toBe(true);
      expect(fontAt(workbook, MAX_ROW, MAX_COL).size).toBe(11);
      expect(sheet.instance.redo()).toBe(true);
      expect(fontAt(workbook, MAX_ROW, MAX_COL).size).toBe(12);
    },
  );

  it('does not read a baseline for a font clear that creates no override', async () => {
    const fixture = await mountFontToolbar(0);
    mounted.push(fixture);
    mutators.selectAll(fixture.sheet.instance.store);
    const snapshot = vi.spyOn(fixture.workbook, 'getSheetFontAxisSnapshot');

    expect(
      setFont(fixture.sheet.instance.store.getState(), fixture.sheet.instance.store, {
        fontSize: null,
      }),
    ).toBe(true);
    expect(snapshot).not.toHaveBeenCalled();
    expect(fixture.sheet.instance.store.getState().format.sheetFontOverrides?.size ?? 0).toBe(0);
  });

  it('keeps the old fallback when baseline capture returns null', async () => {
    const fixture = await mountFontToolbar(1);
    mounted.push(fixture);
    mutators.selectAll(fixture.sheet.instance.store);
    const snapshot = vi.spyOn(fixture.workbook, 'getSheetFontAxisSnapshot').mockReturnValue(null);

    expect(() =>
      setFont(fixture.sheet.instance.store.getState(), fixture.sheet.instance.store, {
        fontSize: 24,
      }),
    ).not.toThrow();
    expect(
      fixture.sheet.instance.store.getState().format.sheetFontOverrides?.get(0)?.baseline,
    ).toBe(undefined);
    expect(snapshot).toHaveBeenCalled();
  });

  it('preserves a baseline already carried by an existing override', async () => {
    const fixture = await mountFontToolbar(1);
    mounted.push(fixture);
    const { workbook, sheet } = fixture;
    const baseline = workbook.getSheetFontAxisSnapshot(0);
    if (!baseline) throw new Error('missing native baseline');
    const storedBaseline: SheetFontAxisSnapshot = {
      columns: baseline.columns.map((entry) => ({ ...entry })),
      rows: baseline.rows.map((entry) => ({ ...entry })),
    };
    sheet.instance.store.setState((state) => ({
      ...state,
      format: {
        ...state.format,
        sheetFontOverrides: new Map([[0, { patch: { fontSize: 18 }, baseline: storedBaseline }]]),
      },
    }));
    mutators.selectAll(sheet.instance.store);
    const snapshot = vi.spyOn(workbook, 'getSheetFontAxisSnapshot');

    expect(setFont(sheet.instance.store.getState(), sheet.instance.store, { fontSize: 24 })).toBe(
      true,
    );
    expect(snapshot).not.toHaveBeenCalled();
    expect(sheet.instance.store.getState().format.sheetFontOverrides?.get(0)?.baseline).toEqual(
      storedBaseline,
    );
  });
});
