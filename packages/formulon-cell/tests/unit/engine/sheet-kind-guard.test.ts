import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { checkAccessibility } from '../../../src/commands/accessibility-check.js';
import { insertCells } from '../../../src/commands/cell-shift.js';
import { circularReferenceEntries } from '../../../src/commands/circular-references.js';
import { copy } from '../../../src/commands/clipboard/copy.js';
import { insertCopiedBand } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { captureSnapshotFromCopyResult } from '../../../src/commands/clipboard/snapshot.js';
import { collectAllFormulas } from '../../../src/commands/formula-records.js';
import { History } from '../../../src/commands/history.js';
import { applyPrecisionAsDisplayed } from '../../../src/commands/precision-as-displayed.js';
import { writeWorkbook } from '../../../src/commands/workbook-save.js';
import { pageSetupPatchFromEngine } from '../../../src/engine/print-sync.js';
import { SheetKind, SheetVisibility, type Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { workbookTables } from '../../../src/engine/workbook-handle-tables.js';
import { collectWorkbookStatistics } from '../../../src/interact/mac-workbook-statistics.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { refuseNonWorksheetCalls, sheetKindWorkbook } from './sheet-kind-workbook.js';

/** Switch tabs the way a click on the tab does. */
const clickTab = (host: HTMLElement, sheet: number): void => {
  const tab = host.querySelector<HTMLElement>(
    `.fc-host__sheetbar-tab[data-fc-sheet-index="${sheet}"]`,
  );
  if (!tab) throw new Error(`no tab for sheet ${sheet}`);
  tab.click();
};

const chartsheetFixture = (): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, '../../fixtures/charts/chart_chartsheet.xlsx')));

/** The data worksheet, a macrosheet and a dialogsheet; the engine refuses what they do not answer. */
const kindsWorkbook = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.loadBytes(sheetKindWorkbook());
  refuseNonWorksheetCalls(wb);
  return wb;
};

describe('a workbook holding sheets that are not worksheets', () => {
  const opened: WorkbookHandle[] = [];
  let mounted: MountedStubSheet | null = null;
  afterEach(() => {
    mounted?.dispose();
    mounted = null;
    for (const wb of opened.splice(0)) wb.dispose();
  });
  const open = async (): Promise<WorkbookHandle> => {
    const wb = await kindsWorkbook();
    opened.push(wb);
    return wb;
  };

  it('reports each sheet type', async () => {
    const wb = await open();
    expect([0, 1, 2].map((i) => wb.sheetKind(i))).toEqual([
      SheetKind.Worksheet,
      SheetKind.Macrosheet,
      SheetKind.Dialogsheet,
    ]);
  });

  it('reads no cells, tables or view from a sheet that is not a worksheet', async () => {
    const wb = await open();
    expect([...wb.cells(0)]).toHaveLength(1);
    for (const sheet of [1, 2]) {
      expect([...wb.cells(sheet)]).toEqual([]);
      expect([...wb.physicalCells(sheet)]).toEqual([]);
      expect(wb.getSheetView(sheet)).toBeNull();
      expect(wb.getSheetPr(sheet)).toBeNull();
      expect(wb.paginateSheet(sheet)).toBeNull();
      expect(wb.sheetVisibility(sheet)).toBe(SheetVisibility.Visible);
    }
    expect(workbookTables((wb as unknown as { wb: Workbook }).wb)).toEqual([]);
  });

  it('keeps the worksheet page setup and reads none from the others', async () => {
    const wb = await open();
    expect(() => pageSetupPatchFromEngine(wb, 0)).not.toThrow();
    for (const sheet of [1, 2]) expect(pageSetupPatchFromEngine(wb, sheet)).toEqual({});
  });

  it('runs the workbook-wide sweeps over the worksheet alone', async () => {
    const wb = await open();
    const store = createSpreadsheetStore();
    expect(collectWorkbookStatistics(wb, store).populatedCells).toBe(1);
    expect(collectAllFormulas(wb)).toEqual([]);
    expect(applyPrecisionAsDisplayed(wb, store, null)).toBe(0);
    expect(circularReferenceEntries(wb, 0)).toEqual([]);
    expect(circularReferenceEntries(wb, 1)).toEqual([]);
    mutators.replaceCells(store, wb.cells(0));
    expect(() => checkAccessibility(store.getState(), wb)).not.toThrow();
    expect(() => writeWorkbook(wb, 'xlsx')).not.toThrow();
  });

  it('shifts cells and inserts a cut band without reading the other sheets', async () => {
    const wb = await open();
    const store = createSpreadsheetStore();
    wb.attachStore(store);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    wb.setNumber({ sheet: 0, row: 2, col: 0 }, 3);
    const cell = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    expect(insertCells(store, wb, null, cell, 'down')).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 7 });

    mutators.replaceCells(store, wb.cells(0));
    const source = { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 16_383 };
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    mutators.setRange(store, source);
    const result = copy(store.getState());
    if (!result) throw new Error('missing copy result');
    const snapshot = captureSnapshotFromCopyResult(store.getState(), result, 'cut');
    mutators.setCopyRange(store, source, 'cut');
    const history = new History();
    wb.attachHistory(history);
    const moved = insertCopiedBand(store, wb, history, snapshot, { sheet: 0, row: 4, col: 0 });
    expect(moved).toMatchObject({ status: 'applied' });
  });

  it('mounts on a macrosheet and a dialogsheet and switches tabs', async () => {
    const wb = await open();
    mounted = await mountStubSheet({ workbook: wb });
    const { instance, host } = mounted;
    for (const sheet of [1, 2, 0, 2, 1, 0]) {
      clickTab(host, sheet);
      expect(instance.store.getState().data.sheetIndex).toBe(sheet);
    }
    const layout = instance.store.getState().layout;
    expect(layout.sheetKinds?.get(1)).toBe(SheetKind.Macrosheet);
    expect(layout.sheetKinds?.get(2)).toBe(SheetKind.Dialogsheet);
    expect(instance.store.getState().data.cells.size).toBe(1);
    clickTab(host, 1);
    expect(instance.store.getState().data.cells.size).toBe(0);
    expect(() => instance.workbook.save()).not.toThrow();
  });

  it('mounts with a macrosheet active', async () => {
    const wb = await open();
    expect(wb.setActiveSheet(1)).toBe(true);
    mounted = await mountStubSheet({ workbook: wb });
    expect(mounted.instance.store.getState().data.sheetIndex).toBe(1);
  });
});

describe('a workbook holding a chartsheet', () => {
  let mounted: MountedStubSheet | null = null;
  afterEach(() => {
    mounted?.dispose();
    mounted = null;
  });

  it('mounts on the chartsheet, switches tabs and saves when the engine refuses worksheet calls', async () => {
    const wb = await WorkbookHandle.loadBytes(chartsheetFixture());
    refuseNonWorksheetCalls(wb);
    mounted = await mountStubSheet({ workbook: wb });
    const { instance, host } = mounted;
    expect(host.dataset.fcSheetKind).toBe('chartsheet');
    expect(instance.store.getState().data.cells.size).toBe(0);
    clickTab(host, 1);
    clickTab(host, 0);
    expect(instance.store.getState().charts.charts).toHaveLength(1);
    expect(() => instance.workbook.save()).not.toThrow();
    expect(() => collectWorkbookStatistics(wb, instance.store)).not.toThrow();
    expect(() => checkAccessibility(instance.store.getState(), wb)).not.toThrow();
    expect(pageSetupPatchFromEngine(wb, 0)).not.toHaveProperty('printArea');
  });
});
