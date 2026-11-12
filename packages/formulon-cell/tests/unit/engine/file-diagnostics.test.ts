import { FormulonError, StatusCode, WorkbookFormat } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import {
  emptyLoadDiagnostics,
  emptySaveDiagnostics,
  fileLossEntries,
  hasLoadLosses,
  loadDiagnosticsFromEngine,
  saveDiagnosticsFromEngine,
} from '../../../src/engine/file-diagnostics.js';
import type { Formulon, Workbook } from '../../../src/engine/types.js';
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
});
