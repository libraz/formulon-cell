import { describe, expect, it, vi } from 'vitest';

import type { EvalResult, FormulonModule, Status, Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const ok: Status = { ok: true, status: 0, message: '', context: '' };
const failed = (message: string): Status => ({ ok: false, status: 1, message, context: '' });

const numberResult = (number: number): EvalResult => ({
  status: ok,
  value: { kind: 1, number, boolean: 0, text: '', errorCode: 0 },
});

const textResult = (text: string): EvalResult => ({
  status: ok,
  value: { kind: 3, number: 0, boolean: 0, text, errorCode: 0 },
});

type RawWorkbook = Record<string, unknown> & {
  delete: ReturnType<typeof vi.fn>;
};

type FakeEngine = {
  handle: WorkbookHandle;
  active: RawWorkbook;
  fresh: RawWorkbook[];
  evalFormula: ReturnType<typeof vi.fn>;
  createDefault: ReturnType<typeof vi.fn>;
};

const makeEngine = (
  opts: {
    profile?: string | null;
    profileStatus?: Status;
    freshSet?: (profile: string) => Status;
    freshEvaluate?: (sheet: number, row: number, col: number, formula: string) => EvalResult;
    activeEvaluate?: ReturnType<typeof vi.fn>;
    includeLegacyProfileApi?: boolean;
  } = {},
): FakeEngine => {
  const fresh: RawWorkbook[] = [];
  const evalFormula = vi.fn((formula: string) => textResult(`legacy:${formula}`));
  const createDefault = vi.fn(() => {
    const next: RawWorkbook = {
      delete: vi.fn(),
      setExcelProfileId: vi.fn((profile: string) => opts.freshSet?.(profile) ?? ok),
      evaluateFormulaText: vi.fn(
        (sheet: number, row: number, col: number, formula: string) =>
          opts.freshEvaluate?.(sheet, row, col, formula) ?? numberResult(42),
      ),
    };
    fresh.push(next);
    return next;
  });
  const active: RawWorkbook = {
    delete: vi.fn(),
    evaluateFormulaText: opts.activeEvaluate ?? vi.fn(() => numberResult(999)),
  };
  if (opts.includeLegacyProfileApi !== false) {
    active.excelProfileId = vi.fn(() => ({
      status: opts.profileStatus ?? ok,
      value: opts.profile,
    }));
    active.setExcelProfileId = vi.fn(() => ok);
  }
  const module = {
    Workbook: { createDefault },
    evalFormula,
    versionString: () => 'test',
  } as unknown as FormulonModule;
  const Ctor = WorkbookHandle as unknown as new (
    module: FormulonModule,
    wb: Workbook,
  ) => WorkbookHandle;
  return {
    handle: new Ctor(module, active as unknown as Workbook),
    active,
    fresh,
    evalFormula,
    createDefault,
  };
};

describe('WorkbookHandle.evalFormula profile isolation', () => {
  it('evaluates against a fresh workbook with the active raw profile', () => {
    const activeEvaluate = vi.fn(() => numberResult(999));
    const engine = makeEngine({ profile: 'mac-365-future_XY', activeEvaluate });

    const result = engine.handle.evalFormula('=A1');

    expect(result).toEqual(numberResult(42));
    expect(engine.evalFormula).not.toHaveBeenCalled();
    expect(activeEvaluate).not.toHaveBeenCalled();
    expect(engine.createDefault).toHaveBeenCalledOnce();
    expect(engine.fresh[0]?.setExcelProfileId).toHaveBeenCalledWith('mac-365-future_XY');
    expect(engine.fresh[0]?.evaluateFormulaText).toHaveBeenCalledWith(0, 0, 0, '=A1');
    expect(engine.fresh[0]?.delete).toHaveBeenCalledOnce();
  });

  it('keeps the legacy module evaluator when the profile-aware API is absent', () => {
    const engine = makeEngine({ includeLegacyProfileApi: false });

    expect(engine.handle.evalFormula('=1+2')).toEqual(textResult('legacy:=1+2'));
    expect(engine.evalFormula).toHaveBeenCalledWith('=1+2');
    expect(engine.createDefault).not.toHaveBeenCalled();
  });

  it('falls back when the active getter returns no profile', () => {
    const engine = makeEngine({ profile: null });

    expect(engine.handle.evalFormula('=1+2')).toEqual(textResult('legacy:=1+2'));
    expect(engine.evalFormula).toHaveBeenCalledWith('=1+2');
    expect(engine.createDefault).not.toHaveBeenCalled();
  });

  it('propagates an active profile getter failure', () => {
    const engine = makeEngine({
      profile: null,
      profileStatus: failed('profile read failed'),
    });

    expect(() => engine.handle.evalFormula('=1+2')).toThrow('excelProfileId: profile read failed');
    expect(engine.createDefault).not.toHaveBeenCalled();
  });

  it('deletes the fresh workbook when evaluation throws', () => {
    const engine = makeEngine({
      profile: 'win-365-en_US',
      freshEvaluate: () => {
        throw new Error('evaluation exploded');
      },
    });

    expect(() => engine.handle.evalFormula('=1+2')).toThrow('evaluation exploded');
    expect(engine.fresh[0]?.delete).toHaveBeenCalledOnce();
  });

  it('deletes the fresh workbook and rejects a failed profile setter', () => {
    const engine = makeEngine({
      profile: 'win-365-en_US',
      freshSet: () => failed('profile write failed'),
    });

    expect(() => engine.handle.evalFormula('=1+2')).toThrow(
      'setExcelProfileId: profile write failed',
    );
    expect(engine.fresh[0]?.delete).toHaveBeenCalledOnce();
    expect(engine.fresh[0]?.evaluateFormulaText).not.toHaveBeenCalled();
  });

  it('uses the active profile for real Mac and Windows collation', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId('mac-en_US')).toBe(true);
      expect(workbook.evalFormula('="ガ"="ｶﾞ"').value.boolean).toBe(1);
      expect(workbook.setSpreadsheetProfileId('windows-en_US')).toBe(true);
      expect(workbook.evalFormula('="ガ"="ｶﾞ"').value.boolean).toBe(0);
    } finally {
      workbook.dispose();
    }
  });

  it('uses the Japanese profile for localized currency while keeping active cells invisible', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId('windows-ja_JP')).toBe(true);
      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 99);

      const currency = workbook.evalFormula('=YEN(1234.5)');
      expect(currency.value).toMatchObject({ kind: 3, text: '¥1,235' });
      expect(workbook.evalFormula('=A1').value).toMatchObject({ kind: 1, number: 0 });
    } finally {
      workbook.dispose();
    }
  });
});
