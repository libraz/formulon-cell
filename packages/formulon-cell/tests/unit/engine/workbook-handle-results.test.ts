import { FormulonError, StatusCode } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { upsertDefinedName } from '../../../src/commands/named-ranges.js';
import { loadFormulon } from '../../../src/engine/loader.js';
import type { Formulon, Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const failed = (message: string): FormulonError =>
  new FormulonError(StatusCode.InvalidArgument, message);

const throwing = (message: string) => (): never => {
  throw failed(message);
};

const Ctor = WorkbookHandle as unknown as new (module: Formulon, wb: Workbook) => WorkbookHandle;

const makeHandle = (
  raw: Record<string, unknown>,
  module: Record<string, unknown> = {},
): WorkbookHandle =>
  new Ctor(
    // A spread would drop a Proxy surface's trapped members; pass a complete module through as is.
    ('getVersion' in module
      ? module
      : { getVersion: () => 'test', ...module }) as unknown as Formulon,
    // Every engine has a clock; a fake that leaves it out reads as unpinned.
    new Proxy(raw, {
      get: (target, key) =>
        key === 'getPinnedNow' && !(key in target) ? () => null : Reflect.get(target, key),
    }) as unknown as Workbook,
  );

/** An object whose every method fails with a `FormulonError` naming it. */
const failingSurface = (overrides: Record<string, unknown> = {}): Record<string, unknown> =>
  new Proxy(overrides, {
    get: (target, key) =>
      key in target ? target[key as string] : throwing(`${String(key)} failed`),
  });

const metadata = (name: string, availability: number) => ({
  name,
  minArity: 1,
  maxArity: null,
  availability,
  signature: null,
  description: null,
});

describe('WorkbookHandle engine failures', () => {
  it.each([0, 2, 3, 99])('preserves raw function availability %i', (availability) => {
    const wb = makeHandle(
      {},
      {
        listFunctionNames: () => ['SUM'],
        getFunctionMetadata: () => metadata('SUM', availability),
      },
    );
    expect(wb.functionMetadata('SUM')?.availability).toBe(availability);
  });

  it('keeps the engine availability and prevents provider overrides', () => {
    const wb = makeHandle(
      {},
      {
        listFunctionNames: () => ['CUBEVALUE'],
        getFunctionMetadata: () => metadata('CUBEVALUE', 3),
      },
    );
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

  it('keeps the empty-string lookup fallback when a function name does not match', async () => {
    const wb = makeHandle(
      { getProfileId: () => 'win-365-en_US' },
      {
        localizeFunctionName: throwing('unknown function'),
        canonicalizeFunctionName: throwing('unknown function'),
      },
    );
    expect(wb.localizeFunctionName('UNKNOWN_FUNCTION')).toBe('');
    expect(wb.canonicalizeFunctionName('UNKNOWN_FUNCTION')).toBe('');
    const stub = await loadFormulon({ preferStub: true });
    const stubWb = new Ctor(stub, stub.Workbook.createDefault());
    expect(stubWb.localizeFunctionName('SUM')).toBeNull();
    expect(stubWb.canonicalizeFunctionName('SUM')).toBeNull();
  });

  it('returns plain arrays the caller may keep', () => {
    const rawNames = Object.freeze(['USDOLLAR']);
    const rawFormats = Object.freeze([]);
    const wb = makeHandle(
      { sheet: () => ({ listCfRules: () => rawFormats }) },
      { listFunctionNames: () => rawNames, getFunctionMetadata: () => null },
    );

    const names = wb.functionNames();
    const formats = wb.getConditionalFormats(0);
    expect(names).toEqual(['USDOLLAR']);
    expect(names).not.toBe(rawNames);
    expect(formats).not.toBe(rawFormats);
  });

  it('does not turn a failed required count into a valid zero', () => {
    const wb = makeHandle({
      get sheetCount(): number {
        throw failed('sheet count failed');
      },
    });

    expect(() => wb.sheetCount).toThrow('sheet count failed');
  });

  it('does not silently skip cells or defined names when their list read fails', () => {
    const cellWb = makeHandle({
      sheet: () => ({ listCells: throwing('cell count failed') }),
    });
    const nameWb = makeHandle({ listDefinedNames: throwing('defined-name count failed') });

    expect(() => [...cellWb.physicalCells(0)]).toThrow('cell count failed');
    expect(() => [...nameWb.definedNames()]).toThrow('defined-name count failed');
  });

  it('keeps optional failed list reads on their existing fallbacks', () => {
    const sheet = failingSurface({ ooxml: failingSurface() });
    const wb = makeHandle(
      failingSurface({
        sheetCount: 1,
        hasSheet: () => true,
        sheet: () => sheet,
        ooxml: failingSurface(),
      }),
      failingSurface({ getVersion: () => 'test' }),
    );

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

  it('suppresses automatic recalc in manual mode and recalcs when the mode read fails', () => {
    let recalcCalls = 0;
    let modeFails = false;
    const wb = makeHandle({
      getCalcMode: () => {
        if (modeFails) throw failed('calc mode failed');
        return 1;
      },
      setCalcMode: () => {},
      recalc: () => {
        recalcCalls += 1;
      },
    });

    wb.recalcAuto();
    expect(recalcCalls).toBe(0);

    modeFails = true;
    expect(wb.calcMode()).toBeNull();
    wb.recalcAuto();
    expect(recalcCalls).toBe(1);
  });
});

describe('WorkbookHandle defined-name key resolution', () => {
  it('surfaces a real engine refusal instead of falling back to the typed key', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const key = wb.resolveDefinedNameKey('Broken', '=(', -1);
      expect(key.ok).toBe(false);
      if (key.ok) return;
      expect(key.status).not.toBeNull();
      expect(key.status).not.toBe(StatusCode.NotImplemented);

      const upsert = upsertDefinedName(wb, 'Broken', '=(', -1);
      expect(upsert).toEqual({ ok: false, reason: 'engine-failed', status: key.status });
      expect([...wb.definedNames()]).toHaveLength(0);
    } finally {
      wb.dispose();
    }
  });

  it('falls back to the typed key only for a stub resolver', () => {
    const wb = makeHandle({
      resolveDefinedNameKey: () => {
        throw new FormulonError(StatusCode.NotImplemented, 'stub');
      },
    });
    Object.defineProperty(wb, 'capabilities', {
      value: { ...wb.capabilities, definedNameMutate: true },
    });
    expect(wb.resolveDefinedNameKey('print_area', '=Sheet1!$A$1', -1)).toEqual({
      ok: true,
      name: 'print_area',
      localSheetId: -1,
    });
  });
});
