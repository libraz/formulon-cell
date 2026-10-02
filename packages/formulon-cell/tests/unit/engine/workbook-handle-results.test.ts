import { describe, expect, it } from 'vitest';
import type { FormulonModule, Status, Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const good: Status = { ok: true, status: 0, message: '', context: '' };
const failed = (message: string): Status => ({ ok: false, status: 1, message, context: '' });

const failedList = (message: string): unknown[] & { readonly status: Status } =>
  Object.assign([], { status: failed(message) });

const makeHandle = (raw: Record<string, unknown>): WorkbookHandle => {
  const Ctor = WorkbookHandle as unknown as new (
    module: FormulonModule,
    wb: Workbook,
  ) => WorkbookHandle;
  const module = { versionString: () => 'test' } as unknown as FormulonModule;
  return new Ctor(module, raw as unknown as Workbook);
};

describe('WorkbookHandle 0.12 result envelopes', () => {
  it.each([0, 2, 3, 99])('preserves raw function availability %i', (availability) => {
    const wb = makeHandle({
      functionNames: () => Object.assign(['SUM'], { status: good }),
      functionMetadata: () => ({
        ok: true,
        name: 'SUM',
        minArity: 1,
        maxArity: null,
        availability,
      }),
    });
    expect(wb.functionMetadata('SUM')?.availability).toBe(availability);
  });

  it('keeps availability absent for a legacy engine and prevents provider overrides', () => {
    const legacy = makeHandle({
      functionNames: () => Object.assign(['SUM'], { status: good }),
      functionMetadata: () => ({ ok: true, name: 'SUM', minArity: 1, maxArity: null }),
    });
    expect(Object.hasOwn(legacy.functionMetadata('SUM') ?? {}, 'availability')).toBe(false);
    const wb = makeHandle({
      functionNames: () => Object.assign(['CUBEVALUE'], { status: good }),
      functionMetadata: () => ({
        ok: true,
        name: 'CUBEVALUE',
        minArity: 1,
        maxArity: null,
        availability: 3,
      }),
    });
    const override = { description: 'Host description', availability: 0 };
    wb.setFunctionMetadataProvider({ CUBEVALUE: override });
    expect(wb.functionMetadata('CUBEVALUE')).toMatchObject({
      availability: 3,
      description: 'Host description',
    });
  });

  it('exposes native implemented, environment-bound, and unavailable function states', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.functionMetadata('ACOS')?.availability).toBe(0);
      expect(wb.functionMetadata('INFO')?.availability).toBe(2);
      expect(wb.functionMetadata('CUBEVALUE')?.availability).toBe(3);
      expect(wb.functionMetadata('WEBSERVICE')?.availability).toBe(3);
      expect(wb.functionMetadata('UNKNOWN_FUNCTION')).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('keeps the empty-string lookup fallback when a function name does not match', () => {
    const wb = makeHandle({
      localizeFunctionName: () => ({ status: failed('unknown function'), value: '' }),
      canonicalizeFunctionName: () => ({ status: failed('unknown function'), value: '' }),
    });
    expect(wb.localizeFunctionName('UNKNOWN_FUNCTION')).toBe('');
    expect(wb.canonicalizeFunctionName('UNKNOWN_FUNCTION')).toBe('');
    expect(makeHandle({}).localizeFunctionName('SUM')).toBeNull();
    expect(makeHandle({}).canonicalizeFunctionName('SUM')).toBeNull();
  });

  it('returns plain arrays without exposing raw engine status properties', () => {
    const rawNames = Object.assign(['USDOLLAR'], { status: good });
    const rawFormats = Object.assign([], { status: good });
    const wb = makeHandle({
      functionNames: () => rawNames,
      functionMetadata: () => ({ ok: true }),
      getConditionalFormats: () => rawFormats,
      addConditionalFormat: () => ({ status: good, index: 0 }),
      removeConditionalFormatAt: () => good,
      clearConditionalFormats: () => good,
    });

    const names = wb.functionNames();
    const formats = wb.getConditionalFormats(0);
    expect(names).toEqual(['USDOLLAR']);
    expect(Object.hasOwn(names ?? {}, 'status')).toBe(false);
    expect(Object.hasOwn(formats, 'status')).toBe(false);
    expect(rawNames.status).toBe(good);
    expect(rawFormats.status).toBe(good);
  });

  it('does not turn a failed required count into a valid zero', () => {
    const wb = makeHandle({
      sheetCount: () => ({ status: failed('sheet count failed'), value: 0 }),
    });

    expect(() => wb.sheetCount).toThrow('sheetCount: sheet count failed');
  });

  it('does not silently skip cells or defined names when their count fails', () => {
    const cellWb = makeHandle({
      cellCount: () => ({ status: failed('cell count failed'), value: 0 }),
    });
    const nameWb = makeHandle({
      definedNameCount: () => ({ status: failed('defined-name count failed'), value: 0 }),
    });

    expect(() => [...cellWb.physicalCells(0)]).toThrow('cellCount(0): cell count failed');
    expect(() => [...nameWb.definedNames()]).toThrow('definedNameCount: defined-name count failed');
  });

  it('keeps optional failed list reads on their existing fallbacks', () => {
    const wb = makeHandle({
      getMerges: () => failedList('merges failed'),
      addMerge: () => good,
      removeMerge: () => good,
      clearMerges: () => good,
      getExternalLinks: () => failedList('links failed'),
      getValidations: () => failedList('validations failed'),
      addValidation: () => good,
      clearValidations: () => good,
      removeValidationAt: () => good,
      getHyperlinks: () => failedList('hyperlinks failed'),
      addHyperlink: () => good,
      removeHyperlink: () => good,
      removeHyperlinkAt: () => good,
      clearHyperlinks: () => good,
      getConditionalFormats: () => failedList('conditional formats failed'),
      addConditionalFormat: () => ({ status: good, index: 0 }),
      removeConditionalFormatAt: () => good,
      clearConditionalFormats: () => good,
      functionNames: () => failedList('function names failed'),
      functionMetadata: () => ({ ok: true, name: 'SUM', minArity: 0, maxArity: null }),
      precedents: () => failedList('precedents failed'),
      dependents: () => failedList('dependents failed'),
      spillInfo: () => ({
        status: failed('spill failed'),
        engaged: true,
        anchorRow: 1,
        anchorCol: 1,
        rows: 2,
        cols: 2,
      }),
      excelProfileId: () => ({ status: failed('profile failed'), value: '' }),
      setExcelProfileId: () => good,
      tableCount: () => ({ status: failed('table count failed'), value: 0 }),
      passthroughCount: () => ({ status: failed('passthrough count failed'), value: 0 }),
      cellStyleCount: () => ({ status: failed('cell-style count failed'), value: 0 }),
      cellStyleXfCount: () => ({ status: failed('cell-style xf count failed'), value: 0 }),
      getCellStyle: () => ({ status: failed('cell-style failed') }),
      getCellStyleXf: () => ({ status: failed('cell-style xf failed') }),
      getDxf: () => ({ status: failed('dxf failed') }),
      addDxf: () => ({ status: good, index: 0 }),
      dxfCount: () => ({ status: failed('dxf count failed'), value: 0 }),
    });

    expect(wb.getMerges(0)).toEqual([]);
    expect(wb.getExternalLinks()).toEqual([]);
    expect(wb.getValidationsForSheet(0)).toEqual([]);
    expect(wb.getHyperlinks(0)).toEqual([]);
    expect(wb.getConditionalFormats(0)).toEqual([]);
    expect(wb.functionNames()).toBeNull();
    expect(wb.precedents({ sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(wb.dependents({ sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(wb.spillInfo(0, 0, 0)).toBeNull();
    expect(wb.spreadsheetProfileId()).toBeNull();
    expect(wb.getTables()).toEqual([]);
    expect(wb.getPassthroughs()).toEqual([]);
    expect(wb.cellStyleCount()).toBe(0);
    expect(wb.cellStyleXfCount()).toBe(0);
    expect(wb.dxfCount()).toBe(0);
  });

  it('suppresses automatic recalc in manual mode while preserving failure status', () => {
    let recalcCalls = 0;
    let mode: { status: Status; value: 0 | 1 | 2 } = { status: good, value: 1 };
    const wb = makeHandle({
      calcMode: () => mode,
      setCalcMode: () => good,
      recalc: () => {
        recalcCalls += 1;
        return good;
      },
    });

    wb.recalcAuto();
    expect(recalcCalls).toBe(0);

    mode = { status: failed('calc mode failed'), value: 0 };
    expect(() => wb.recalcAuto()).toThrow('calcMode: calc mode failed');
  });
});
