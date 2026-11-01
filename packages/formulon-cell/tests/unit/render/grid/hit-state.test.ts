import { describe, expect, it } from 'vitest';
import type { CellValue } from '../../../../src/engine/types.js';
import {
  detectErrorKind,
  detectValidationViolation,
  ERROR_TRIANGLE_COLOR,
  gridHitState,
  isPlainTextOverflowCandidate,
  normalizeFormatLocale,
  overflowSides,
  shouldShowValidationChevron,
  VALIDATION_TRIANGLE_COLOR,
} from '../../../../src/render/grid/hit-state.js';

describe('render/grid/hit-state', () => {
  describe('error/validation triangle colors', () => {
    it('exposes distinct sentinel colors for error vs validation triangles', () => {
      expect(ERROR_TRIANGLE_COLOR).toBe('#008000');
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

  describe('shouldShowValidationChevron', () => {
    it('shows only list-validation chevrons whose in-cell dropdown is enabled', () => {
      expect(shouldShowValidationChevron({ kind: 'list', source: ['A', 'B'] })).toBe(true);
      expect(
        shouldShowValidationChevron({ kind: 'list', source: ['A', 'B'], showDropdown: false }),
      ).toBe(false);
      expect(shouldShowValidationChevron({ kind: 'whole', op: 'between', a: 1, b: 10 })).toBe(
        false,
      );
      expect(shouldShowValidationChevron(undefined)).toBe(false);
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

    it('rejects wrapped, rotated, shrunk and fill/justify-style text', () => {
      expect(isPlainTextOverflowCandidate({ ...baseInput, format: { wrap: true } })).toBe(false);
      expect(isPlainTextOverflowCandidate({ ...baseInput, format: { rotation: 45 } })).toBe(false);
      expect(isPlainTextOverflowCandidate({ ...baseInput, format: { shrinkToFit: true } })).toBe(
        false,
      );
      for (const align of ['fill', 'justify', 'distributed', 'centerContinuous'] as const) {
        expect(isPlainTextOverflowCandidate({ ...baseInput, format: { align } })).toBe(false);
      }
    });

    it('accepts left, right, centred, indented and formula-result text', () => {
      for (const align of ['left', 'right', 'center'] as const) {
        expect(isPlainTextOverflowCandidate({ ...baseInput, format: { align, indent: 1 } })).toBe(
          true,
        );
      }
      expect(isPlainTextOverflowCandidate({ ...baseInput, formula: '="long"' })).toBe(true);
    });

    it('spills toward the side the alignment grows', () => {
      expect(overflowSides(undefined, false)).toEqual({ left: false, right: true });
      expect(overflowSides(undefined, true)).toEqual({ left: true, right: false });
      expect(overflowSides('right', false)).toEqual({ left: true, right: false });
      expect(overflowSides('center', true)).toEqual({ left: true, right: true });
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

  describe('gridHitState', () => {
    it('starts empty and returns the same state for the same surface', () => {
      const surface = document.createElement('div');
      const hits = gridHitState(surface);
      expect(hits.fillHandle).toBeNull();
      expect(hits.validationChevron).toBeNull();
      expect(hits.totalsChevron).toBeNull();
      expect(hits.errorTriangles).toEqual([]);
      expect(gridHitState(surface)).toBe(hits);
    });

    it('keeps each surface separate', () => {
      const a = document.createElement('div');
      const b = document.createElement('div');
      gridHitState(a).fillHandle = { x: 10, y: 20, w: 6, h: 6 };
      expect(gridHitState(b).fillHandle).toBeNull();
    });
  });
});
