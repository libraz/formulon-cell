import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { FormulonError, StatusCode, WorkbookFormat } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import {
  emptyLoadDiagnostics,
  emptySaveDiagnostics,
  fileLossEntries,
  hasLoadLosses,
  loadDiagnosticsFromEngine,
  saveDiagnosticsFromEngine,
} from '../../../src/engine/file-diagnostics.js';
import { type Formulon, PivotAggregation, type Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { buildFileLossReport } from '../../../src/interact/spreadsheet-compatibility-report.js';

const makeHandle = (raw: Record<string, unknown>): WorkbookHandle => {
  const Ctor = WorkbookHandle as unknown as new (module: Formulon, wb: Workbook) => WorkbookHandle;
  const module = { getVersion: () => 'test' } as unknown as Formulon;
  return new Ctor(module, {
    getCalcMode: () => 0,
    getPinnedNow: () => null,
    ...raw,
  } as unknown as Workbook);
};

describe('file loss entries', () => {
  it('is empty for zero counters', () => {
    expect(fileLossEntries(emptyLoadDiagnostics(), emptySaveDiagnostics())).toEqual([]);
    expect(hasLoadLosses(emptyLoadDiagnostics())).toBe(false);
  });

  it('maps each non-zero counter to its own row, load before save', () => {
    const load = loadDiagnosticsFromEngine({ undecodedFormulaCount: 3, skippedFeatureCount: 2 });
    const save = saveDiagnosticsFromEngine({ droppedPartCount: 1, droppedRelationshipCount: 1 });
    expect(fileLossEntries(load, save)).toEqual([
      { id: 'undecodedFormula', phase: 'load', count: 3 },
      { id: 'skippedFeature', phase: 'load', count: 2 },
      { id: 'droppedPart', phase: 'save', count: 1 },
      { id: 'droppedRelationship', phase: 'save', count: 1 },
    ]);
    expect(hasLoadLosses(load)).toBe(true);
  });

  it('coerces missing or invalid counters to zero', () => {
    expect(saveDiagnosticsFromEngine({ droppedPartCount: Number.NaN })).toEqual(
      emptySaveDiagnostics(),
    );
  });
});

describe('WorkbookHandle diagnostics wrappers', () => {
  it('save() routes through saveWithDiagnostics and caches the counters', () => {
    const formats: unknown[] = [];
    const wb = makeHandle({
      saveWithDiagnostics: (format: unknown) => {
        formats.push(format);
        return {
          bytes: new Uint8Array([1, 2]),
          downgradedFormulaCount: 4,
          deferredFeatureCount: 0,
          droppedPartCount: 0,
          droppedRelationshipCount: 0,
          renumberedPartCount: 0,
        };
      },
    });
    expect(wb.lastSaveDiagnostics().downgradedFormulaCount).toBe(0);
    expect(Array.from(wb.save())).toEqual([1, 2]);
    expect(formats).toEqual([WorkbookFormat.Xlsx]);
    expect(wb.lastSaveDiagnostics().downgradedFormulaCount).toBe(4);
    expect(wb.saveWithDiagnostics('xlsb').diagnostics.downgradedFormulaCount).toBe(4);
    expect(formats).toEqual([WorkbookFormat.Xlsx, WorkbookFormat.Xlsb]);
  });

  it('reports zero counters for a save without losses', () => {
    const wb = makeHandle({
      saveWithDiagnostics: () => ({
        bytes: new Uint8Array([9]),
        downgradedFormulaCount: 0,
        deferredFeatureCount: 0,
        droppedPartCount: 0,
        droppedRelationshipCount: 0,
        renumberedPartCount: 0,
      }),
    });
    expect(wb.saveWithDiagnostics().diagnostics).toEqual(emptySaveDiagnostics());
  });

  it('throws when the engine save fails', () => {
    const wb = makeHandle({
      saveWithDiagnostics: () => {
        throw new FormulonError(StatusCode.InvalidArgument, 'boom');
      },
    });
    expect(() => wb.save()).toThrow('boom');
  });

  it('reports empty load diagnostics for handles not loaded from bytes', () => {
    expect(makeHandle({}).loadDiagnostics()).toEqual(emptyLoadDiagnostics());
  });
});

describe('buildFileLossReport', () => {
  const wb = makeHandle({});

  it('is empty without losses', () => {
    expect(buildFileLossReport(wb, en.workbookObjects)).toEqual([]);
  });

  it('renders localized rows with occurrence counts', () => {
    const lossy = makeHandle({
      saveWithDiagnostics: () => ({
        bytes: new Uint8Array(1),
        downgradedFormulaCount: 7,
      }),
    });
    lossy.save();
    expect(buildFileLossReport(lossy, en.workbookObjects)).toEqual([
      {
        severity: 'warning',
        label: 'Formulas saved as values only · On save',
        detail: '7 occurrences',
      },
    ]);
    expect(buildFileLossReport(lossy, ja.workbookObjects)[0]?.detail).toBe('7 件');
  });
});

describe('file diagnostics on the real engine', () => {
  it('reads zero losses after a clean save and reload', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    const { bytes, diagnostics } = wb.saveWithDiagnostics();
    expect(diagnostics).toEqual(emptySaveDiagnostics());
    const reloaded = await WorkbookHandle.loadBytes(bytes);
    expect(reloaded.loadDiagnostics()).toEqual(emptyLoadDiagnostics());
    expect(hasLoadLosses(reloaded.loadDiagnostics())).toBe(false);
    wb.dispose();
    reloaded.dispose();
  });

  // A PivotTable the engine models is written from its model, so an .xlsb save drops none of its parts.
  const pivotPlacement = (wb: WorkbookHandle) =>
    wb.getPivotTables().map((p) => [p.sheetIndex, p.top, p.left, p.rows, p.cols]);

  it('drops no part when a refreshed .xlsb PivotTable is saved back to .xlsb', async () => {
    const fixture = join(__dirname, '../../fixtures/pivot-excel/xlsb_pivot_base.xlsb');
    const wb = await WorkbookHandle.loadBytes(new Uint8Array(readFileSync(fixture)));
    const placement = pivotPlacement(wb);
    expect(placement).toHaveLength(1);
    for (const id of wb.pivotCacheIds()) expect(wb.refreshPivotCache(id)).toBeNull();
    const { bytes, diagnostics } = wb.saveWithDiagnostics('xlsb');
    expect(diagnostics.droppedPartCount).toBe(0);
    const reloaded = await WorkbookHandle.loadBytes(bytes);
    expect(pivotPlacement(reloaded)).toEqual(placement);
    wb.dispose();
    reloaded.dispose();
  });

  // The report's cells, so a reload can be compared value by value.
  const pivotReport = (wb: WorkbookHandle) =>
    Array.from(wb.pivotCells(0), (c) => [c.addr.row, c.addr.col, c.value]);

  it('writes a PivotTable built in the session to .xlsb, blank source cells included', async () => {
    const wb = await WorkbookHandle.createDefault();
    const rows: (string | number | null)[][] = [
      ['Region', 'Amt'],
      ['East', 5],
      [null, 3],
      ['West', 7],
      ['West', null],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        const addr = { sheet: 0, row: r, col: 3 + c };
        if (typeof v === 'number') wb.setNumber(addr, v);
        else if (v !== null) wb.setText(addr, v);
      }
    }
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 3, r1: 4, c1: 4 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      valueField: 'Amt',
      aggregation: PivotAggregation.Sum,
    });
    expect(created.ok).toBe(true);
    const { bytes, diagnostics } = wb.saveWithDiagnostics('xlsb');
    expect(diagnostics).toEqual(emptySaveDiagnostics());
    const reloaded = await WorkbookHandle.loadBytes(bytes);
    expect(pivotPlacement(reloaded)).toEqual(pivotPlacement(wb));
    expect(pivotReport(wb)).toHaveLength(10);
    expect(pivotReport(reloaded)).toEqual(pivotReport(wb));
    wb.dispose();
    reloaded.dispose();
  });

  it('reports no loss when an unedited .xlsb is saved again, to .xlsb or .xlsx', async () => {
    const fixture = join(__dirname, '../../fixtures/pivot-excel/xlsb_pivot_base.xlsb');
    const wb = await WorkbookHandle.loadBytes(new Uint8Array(readFileSync(fixture)));
    expect(wb.saveWithDiagnostics('xlsb').diagnostics).toEqual(emptySaveDiagnostics());
    expect(wb.saveWithDiagnostics('xlsx').diagnostics).toEqual(emptySaveDiagnostics());
    wb.dispose();
  });
});
