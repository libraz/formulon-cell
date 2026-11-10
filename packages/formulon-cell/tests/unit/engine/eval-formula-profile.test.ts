import { FormulonError, StatusCode, ValueKind } from '@libraz/formulon';
import { describe, expect, it, vi } from 'vitest';
import { loadFormulon } from '../../../src/engine/loader.js';
import type { Formulon, Value, Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const failed = (message: string): FormulonError =>
  new FormulonError(StatusCode.InvalidArgument, message);

const numberResult = (number: number): Value => ({ kind: ValueKind.Number, number });

const textResult = (text: string): Value => ({ kind: ValueKind.Text, text });

type RawWorkbook = Record<string, unknown> & {
  dispose: ReturnType<typeof vi.fn>;
};

type FreshWorkbook = RawWorkbook & {
  setProfileId: ReturnType<typeof vi.fn>;
  evaluateTypedFormula: ReturnType<typeof vi.fn>;
};

type FakeEngine = {
  handle: WorkbookHandle;
  active: RawWorkbook;
  fresh: FreshWorkbook[];
  evalFormula: ReturnType<typeof vi.fn>;
  createDefault: ReturnType<typeof vi.fn>;
};

const makeEngine = (
  opts: {
    profile?: string | null;
    profileError?: FormulonError;
    freshSet?: (profile: string) => void;
    freshEvaluate?: (row: number, col: number, formula: string) => Value;
    activeEvaluate?: ReturnType<typeof vi.fn>;
  } = {},
): FakeEngine => {
  const fresh: FreshWorkbook[] = [];
  const evalFormula = vi.fn((formula: string) => textResult(`legacy:${formula}`));
  const createDefault = vi.fn(() => {
    const evaluateTypedFormula = vi.fn(
      (row: number, col: number, formula: string) =>
        opts.freshEvaluate?.(row, col, formula) ?? numberResult(42),
    );
    const next: FreshWorkbook = {
      dispose: vi.fn(),
      setProfileId: vi.fn((profile: string) => opts.freshSet?.(profile)),
      evaluateTypedFormula,
      sheet: () => ({ evaluateTypedFormula }),
    };
    fresh.push(next);
    return next;
  });
  const activeEvaluate = opts.activeEvaluate ?? vi.fn(() => numberResult(999));
  const active: RawWorkbook = {
    dispose: vi.fn(),
    sheet: () => ({ evaluateTypedFormula: activeEvaluate }),
    getProfileId: vi.fn(() => {
      if (opts.profileError) throw opts.profileError;
      return opts.profile;
    }),
    setProfileId: vi.fn(),
  };
  const module = {
    Workbook: { createDefault },
    evaluateTypedFormula: evalFormula,
    getVersion: () => 'test',
  } as unknown as Formulon;
  const Ctor = WorkbookHandle as unknown as new (module: Formulon, wb: Workbook) => WorkbookHandle;
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
    expect(engine.fresh[0]?.setProfileId).toHaveBeenCalledWith('mac-365-future_XY');
    expect(engine.fresh[0]?.evaluateTypedFormula).toHaveBeenCalledWith(0, 0, '=A1');
    expect(engine.fresh[0]?.dispose).toHaveBeenCalledOnce();
  });

  it('does not evaluate on the stub engine, which has no formula evaluator', async () => {
    const stub = await loadFormulon({ preferStub: true });
    const Ctor = WorkbookHandle as unknown as new (
      module: Formulon,
      wb: Workbook,
    ) => WorkbookHandle;
    const handle = new Ctor(stub, stub.Workbook.createDefault());
    try {
      expect(handle.evalFormula('=1+2')).toBeNull();
    } finally {
      handle.dispose();
    }
  });

  it('falls back when the active getter returns no profile', () => {
    const engine = makeEngine({ profile: null });

    expect(engine.handle.evalFormula('=1+2')).toEqual(textResult('legacy:=1+2'));
    expect(engine.evalFormula).toHaveBeenCalledWith('=1+2');
    expect(engine.createDefault).not.toHaveBeenCalled();
  });

  it('propagates an active profile getter failure', () => {
    const engine = makeEngine({ profile: null, profileError: failed('profile read failed') });

    expect(() => engine.handle.evalFormula('=1+2')).toThrow('profile read failed');
    expect(engine.createDefault).not.toHaveBeenCalled();
  });

  it('disposes the fresh workbook when evaluation throws', () => {
    const engine = makeEngine({
      profile: 'win-365-en_US',
      freshEvaluate: () => {
        throw new Error('evaluation exploded');
      },
    });

    expect(() => engine.handle.evalFormula('=1+2')).toThrow('evaluation exploded');
    expect(engine.fresh[0]?.dispose).toHaveBeenCalledOnce();
  });

  it('disposes the fresh workbook and rejects a failed profile setter', () => {
    const engine = makeEngine({
      profile: 'win-365-en_US',
      freshSet: () => {
        throw failed('profile write failed');
      },
    });

    expect(() => engine.handle.evalFormula('=1+2')).toThrow('profile write failed');
    expect(engine.fresh[0]?.dispose).toHaveBeenCalledOnce();
    expect(engine.fresh[0]?.evaluateTypedFormula).not.toHaveBeenCalled();
  });

  it('uses the active profile for real Mac and Windows collation', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId('mac-en_US')).toBe(true);
      expect(workbook.evalFormula('="ガ"="ｶﾞ"')).toMatchObject({ boolean: true });
      expect(workbook.setSpreadsheetProfileId('windows-en_US')).toBe(true);
      expect(workbook.evalFormula('="ガ"="ｶﾞ"')).toMatchObject({ boolean: false });
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
      expect(currency).toMatchObject({ kind: ValueKind.Text, text: '¥1,235' });
      expect(workbook.evalFormula('=A1')).toMatchObject({ kind: ValueKind.Number, number: 0 });
    } finally {
      workbook.dispose();
    }
  });
});
