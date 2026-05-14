import { describe, expect, it } from 'vitest';

import { coerceInput } from '../../../src/commands/coerce-input.js';
import { resolveListValues, validateAgainst } from '../../../src/commands/validate.js';

const coerce = (raw: string): ReturnType<typeof coerceInput> => coerceInput(raw);

describe('validateAgainst', () => {
  describe('list', () => {
    const inlineList = {
      kind: 'list' as const,
      source: ['Yes', 'No', 'Maybe'],
      allowBlank: true,
      errorStyle: 'stop' as const,
    };

    it('accepts a value present in the inline list', () => {
      expect(validateAgainst(inlineList, coerce('Yes')).ok).toBe(true);
    });

    it('rejects a value not in the inline list', () => {
      const out = validateAgainst(inlineList, coerce('Perhaps'));
      expect(out.ok).toBe(false);
      if (!out.ok) expect(out.severity).toBe('stop');
    });

    it('range-backed list with no resolver accepts anything (spreadsheet parity)', () => {
      const rangeList = {
        kind: 'list' as const,
        source: { ref: 'Sheet1!A1:A3' },
        allowBlank: true,
        errorStyle: 'warning' as const,
      };
      expect(validateAgainst(rangeList, coerce('anything')).ok).toBe(true);
    });

    it('range-backed list with a resolver matches resolved values', () => {
      const rangeList = {
        kind: 'list' as const,
        source: { ref: 'Sheet1!A1:A3' },
        allowBlank: true,
        errorStyle: 'stop' as const,
      };
      const resolver = () => ['Red', 'Green', 'Blue'];
      expect(validateAgainst(rangeList, coerce('Green'), resolver).ok).toBe(true);
      expect(validateAgainst(rangeList, coerce('Yellow'), resolver).ok).toBe(false);
    });
  });

  describe('whole / decimal', () => {
    it('accepts whole numbers in range', () => {
      const v = {
        kind: 'whole' as const,
        op: 'between' as const,
        a: 1,
        b: 10,
        allowBlank: true,
        errorStyle: 'stop' as const,
      };
      expect(validateAgainst(v, coerce('5')).ok).toBe(true);
    });

    it('rejects non-integers for whole', () => {
      const v = {
        kind: 'whole' as const,
        op: 'between' as const,
        a: 1,
        b: 10,
        allowBlank: true,
        errorStyle: 'stop' as const,
      };
      expect(validateAgainst(v, coerce('3.14')).ok).toBe(false);
    });

    it('rejects numbers outside between bounds', () => {
      const v = {
        kind: 'whole' as const,
        op: 'between' as const,
        a: 1,
        b: 10,
        allowBlank: true,
        errorStyle: 'stop' as const,
      };
      expect(validateAgainst(v, coerce('11')).ok).toBe(false);
      expect(validateAgainst(v, coerce('0')).ok).toBe(false);
    });

    it('decimal accepts fractional values', () => {
      const v = {
        kind: 'decimal' as const,
        op: '>' as const,
        a: 1.5,
        allowBlank: true,
        errorStyle: 'stop' as const,
      };
      expect(validateAgainst(v, coerce('1.6')).ok).toBe(true);
      expect(validateAgainst(v, coerce('1.5')).ok).toBe(false);
    });
  });

  describe('textLength', () => {
    it('counts string length against between bounds', () => {
      const v = {
        kind: 'textLength' as const,
        op: 'between' as const,
        a: 3,
        b: 5,
        allowBlank: false,
        errorStyle: 'stop' as const,
      };
      expect(validateAgainst(v, coerce('cat')).ok).toBe(true);
      expect(validateAgainst(v, coerce('abcdef')).ok).toBe(false);
      expect(validateAgainst(v, coerce('ab')).ok).toBe(false);
    });
  });

  describe('allowBlank', () => {
    const v = {
      kind: 'whole' as const,
      op: 'between' as const,
      a: 1,
      b: 10,
      allowBlank: false,
      errorStyle: 'stop' as const,
    };

    it('rejects blank input when allowBlank=false', () => {
      const out = validateAgainst(v, coerce(''));
      expect(out.ok).toBe(false);
    });

    it('accepts blank input when allowBlank=true (default)', () => {
      const accepting = { ...v, allowBlank: true };
      expect(validateAgainst(accepting, coerce('')).ok).toBe(true);
    });
  });

  it('formula input bypasses validation entirely', () => {
    const v = {
      kind: 'whole' as const,
      op: 'between' as const,
      a: 1,
      b: 10,
      allowBlank: true,
      errorStyle: 'stop' as const,
    };
    expect(validateAgainst(v, coerce('=SUM(A1:A2)')).ok).toBe(true);
  });

  it('custom kind always returns ok (engine evaluates the formula elsewhere)', () => {
    const v = {
      kind: 'custom' as const,
      formula: '=ISNUMBER(A1)',
      allowBlank: true,
      errorStyle: 'stop' as const,
    };
    expect(validateAgainst(v, coerce('123')).ok).toBe(true);
  });

  describe('outcome severity passes through', () => {
    it('warning rule reports warning on reject', () => {
      const v = {
        kind: 'whole' as const,
        op: '=' as const,
        a: 5,
        allowBlank: true,
        errorStyle: 'warning' as const,
      };
      const out = validateAgainst(v, coerce('4'));
      expect(out.ok).toBe(false);
      if (!out.ok) expect(out.severity).toBe('warning');
    });

    it('information rule reports information on reject', () => {
      const v = {
        kind: 'whole' as const,
        op: '=' as const,
        a: 5,
        allowBlank: true,
        errorStyle: 'information' as const,
      };
      const out = validateAgainst(v, coerce('4'));
      expect(out.ok).toBe(false);
      if (!out.ok) expect(out.severity).toBe('information');
    });
  });
});

describe('resolveListValues', () => {
  it('returns inline array verbatim', () => {
    expect(
      resolveListValues({
        kind: 'list',
        source: ['a', 'b', 'c'],
        allowBlank: true,
        errorStyle: 'stop',
      }),
    ).toEqual(['a', 'b', 'c']);
  });

  it('returns the resolver result for range refs', () => {
    expect(
      resolveListValues(
        {
          kind: 'list',
          source: { ref: 'Sheet1!B2:B4' },
          allowBlank: true,
          errorStyle: 'stop',
        },
        () => ['x', 'y', 'z'],
      ),
    ).toEqual(['x', 'y', 'z']);
  });

  it('returns [] when no resolver is provided for a range ref', () => {
    expect(
      resolveListValues({
        kind: 'list',
        source: { ref: 'Sheet1!A1:A3' },
        allowBlank: true,
        errorStyle: 'stop',
      }),
    ).toEqual([]);
  });
});
