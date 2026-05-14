import { beforeEach, describe, expect, it } from 'vitest';
import type { CellValue } from '../../../../src/engine/types.js';
import {
  detectErrorKind,
  detectValidationViolation,
  ERROR_TRIANGLE_COLOR,
  getErrorTriangleHits,
  getFillHandleRect,
  getOutlineToggleHits,
  getValidationChevron,
  isPlainTextOverflowCandidate,
  normalizeFormatLocale,
  setErrorTriangleHits,
  setFillHandleRect,
  setOutlineToggles,
  setValidationChevron,
  VALIDATION_TRIANGLE_COLOR,
} from '../../../../src/render/grid/hit-state.js';

describe('render/grid/hit-state', () => {
  beforeEach(() => {
    // Module-level caches — reset between tests so they don't leak.
    setErrorTriangleHits([]);
    setOutlineToggles([]);
    setFillHandleRect(null);
    setValidationChevron(null);
  });

  describe('error/validation triangle colors', () => {
    it('exposes distinct sentinel colors for error vs validation triangles', () => {
      expect(ERROR_TRIANGLE_COLOR).toBe('#2ea043');
      expect(VALIDATION_TRIANGLE_COLOR).toBe('#d24545');
    });
  });

  describe('normalizeFormatLocale', () => {
    it('expands ja → ja-JP and en → en-US', () => {
      expect(normalizeFormatLocale('ja')).toBe('ja-JP');
      expect(normalizeFormatLocale('en')).toBe('en-US');
    });

    it('passes through full BCP-47 tags', () => {
      expect(normalizeFormatLocale('fr-CA')).toBe('fr-CA');
    });

    it('falls back to en-US for an empty string', () => {
      expect(normalizeFormatLocale('')).toBe('en-US');
    });
  });

  describe('detectErrorKind', () => {
    it('flags engine error values', () => {
      const v: CellValue = { kind: 'error', code: 7, text: '#DIV/0!' };
      expect(detectErrorKind(v)).toBe(true);
    });

    it('flags text values that match the error sentinel set', () => {
      for (const s of ['#DIV/0!', '#NAME?', '#REF!', '#VALUE!', '#N/A']) {
        expect(detectErrorKind({ kind: 'text', value: s })).toBe(true);
      }
    });

    it('does not flag plain text or numbers', () => {
      expect(detectErrorKind({ kind: 'text', value: 'hello' })).toBe(false);
      expect(detectErrorKind({ kind: 'number', value: 1 })).toBe(false);
      expect(detectErrorKind({ kind: 'blank' })).toBe(false);
    });
  });

  describe('detectValidationViolation', () => {
    it('returns false when there is no validation rule', () => {
      expect(detectValidationViolation({ kind: 'number', value: 1 }, undefined)).toBe(false);
    });

    it('returns false when the value already carries an error (triangle would clash)', () => {
      expect(
        detectValidationViolation(
          { kind: 'error', code: 7, text: '#DIV/0!' },
          { kind: 'whole', op: 'between', a: 0, b: 10 },
        ),
      ).toBe(false);
    });

    it('returns true when a number falls outside an integer range rule', () => {
      const out = detectValidationViolation(
        { kind: 'number', value: 99 },
        { kind: 'whole', op: 'between', a: 0, b: 10 },
      );
      expect(out).toBe(true);
    });

    it('returns false when a number is within the rule', () => {
      const out = detectValidationViolation(
        { kind: 'number', value: 5 },
        { kind: 'whole', op: 'between', a: 0, b: 10 },
      );
      expect(out).toBe(false);
    });
  });

  describe('isPlainTextOverflowCandidate', () => {
    const baseInput = {
      value: { kind: 'text', value: 'long string' } as CellValue,
      formula: null as string | null,
      format: undefined,
      showFormulas: false,
      displayOverride: null as string | null,
      tableHeader: false,
      hasIcon: false,
      isMergeAnchor: false,
    };

    it('accepts plain unformatted text', () => {
      expect(isPlainTextOverflowCandidate(baseInput)).toBe(true);
    });

    it('rejects header / icon / merge-anchor cells', () => {
      expect(isPlainTextOverflowCandidate({ ...baseInput, tableHeader: true })).toBe(false);
      expect(isPlainTextOverflowCandidate({ ...baseInput, hasIcon: true })).toBe(false);
      expect(isPlainTextOverflowCandidate({ ...baseInput, isMergeAnchor: true })).toBe(false);
    });

    it('rejects formatted text (wrap / rotation / align)', () => {
      expect(isPlainTextOverflowCandidate({ ...baseInput, format: { wrap: true } })).toBe(false);
      expect(isPlainTextOverflowCandidate({ ...baseInput, format: { rotation: 45 } })).toBe(false);
      expect(isPlainTextOverflowCandidate({ ...baseInput, format: { align: 'center' } })).toBe(
        false,
      );
    });

    it('rejects formula cells when "show formulas" is on', () => {
      expect(
        isPlainTextOverflowCandidate({
          ...baseInput,
          formula: '=A1',
          showFormulas: true,
        }),
      ).toBe(false);
    });

    it('rejects non-text values when no display override', () => {
      expect(
        isPlainTextOverflowCandidate({ ...baseInput, value: { kind: 'number', value: 1 } }),
      ).toBe(false);
    });
  });

  describe('hit-rect caches', () => {
    it('round-trips the fill handle rect', () => {
      const rect = { x: 10, y: 20, w: 6, h: 6 };
      setFillHandleRect(rect);
      expect(getFillHandleRect()).toBe(rect);
      setFillHandleRect(null);
      expect(getFillHandleRect()).toBeNull();
    });

    it('round-trips outline toggle hits', () => {
      setOutlineToggles([
        { axis: 'row', level: 1, i0: 0, i1: 4, rect: { x: 0, y: 0, w: 10, h: 10 } },
      ]);
      expect(getOutlineToggleHits()).toHaveLength(1);
    });

    it('round-trips error triangle hits', () => {
      setErrorTriangleHits([
        {
          rect: { x: 0, y: 0, w: 5, h: 5 },
          addr: { sheet: 0, row: 0, col: 0 },
          kind: 'error',
        },
      ]);
      expect(getErrorTriangleHits()).toHaveLength(1);
    });

    it('round-trips the validation chevron', () => {
      setValidationChevron({ rect: { x: 0, y: 0, w: 12, h: 12 }, row: 2, col: 3 });
      expect(getValidationChevron()).toEqual({
        rect: { x: 0, y: 0, w: 12, h: 12 },
        row: 2,
        col: 3,
      });
    });
  });
});
