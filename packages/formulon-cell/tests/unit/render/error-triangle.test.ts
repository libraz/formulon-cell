import { describe, expect, it, vi } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import {
  detectErrorKind,
  detectValidationViolation,
  ERROR_TRIANGLE_COLOR,
  type ErrorTriangleHit,
  isPlainTextOverflowCandidate,
  VALIDATION_TRIANGLE_COLOR,
} from '../../../src/render/grid.js';
import { paintErrorTriangle, paintValidationTriangle } from '../../../src/render/painters.js';
import type { CellValidation } from '../../../src/store/store.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/* Minimal stub matching the surface paintErrorTriangle / paintValidationTriangle
 * touch. We track save/restore + path primitives + the latest fillStyle so we
 * can assert which color a given paint pass used. */
function makeStubCtx(): {
  ctx: CanvasRenderingContext2D;
  calls: string[];
  fills: string[];
} {
  const calls: string[] = [];
  const fills: string[] = [];
  let fillStyle = '';
  const ctx = {
    save: vi.fn(() => calls.push('save')),
    restore: vi.fn(() => calls.push('restore')),
    beginPath: vi.fn(() => calls.push('beginPath')),
    closePath: vi.fn(() => calls.push('closePath')),
    moveTo: vi.fn(() => calls.push('moveTo')),
    lineTo: vi.fn(() => calls.push('lineTo')),
    fill: vi.fn(() => calls.push('fill')),
    get fillStyle(): string {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
      fills.push(v);
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls, fills };
}

describe('detectErrorKind', () => {
  it('returns true for engine error values', () => {
    expect(detectErrorKind({ kind: 'error', code: 7, text: '#DIV/0!' })).toBe(true);
  });

  it('returns true for known string sentinels (DIV/0, NAME?, N/A, etc.)', () => {
    const sentinels = [
      '#DIV/0!',
      '#NAME?',
      '#REF!',
      '#VALUE!',
      '#NUM!',
      '#N/A',
      '#NULL!',
      '#CIRCULAR!',
    ];
    for (const text of sentinels) {
      expect(detectErrorKind({ kind: 'text', value: text })).toBe(true);
    }
  });

  it('returns false for plain text and numbers', () => {
    expect(detectErrorKind({ kind: 'text', value: 'hello' })).toBe(false);
    expect(detectErrorKind({ kind: 'number', value: 42 })).toBe(false);
    expect(detectErrorKind({ kind: 'blank' })).toBe(false);
  });
});

describe('isPlainTextOverflowCandidate', () => {
  const base = {
    value: { kind: 'text' as const, value: 'long label' },
    formula: null,
    showFormulas: false,
    displayOverride: null,
    tableHeader: false,
    hasIcon: false,
    isMergeAnchor: false,
  };

  it('allows ordinary text to overflow into empty neighbors', () => {
    expect(isPlainTextOverflowCandidate(base)).toBe(true);
  });

  it('keeps wrapped, aligned, formula-display, and table cells clipped', () => {
    expect(isPlainTextOverflowCandidate({ ...base, format: { wrap: true } })).toBe(false);
    expect(isPlainTextOverflowCandidate({ ...base, format: { align: 'center' } })).toBe(false);
    expect(
      isPlainTextOverflowCandidate({
        ...base,
        formula: '=A1',
        showFormulas: true,
      }),
    ).toBe(false);
    expect(isPlainTextOverflowCandidate({ ...base, tableHeader: true })).toBe(false);
  });

  it('does not overflow numbers', () => {
    expect(
      isPlainTextOverflowCandidate({
        ...base,
        value: { kind: 'number', value: 123 },
      }),
    ).toBe(false);
  });
});

describe('detectValidationViolation', () => {
  const wholeRange: CellValidation = {
    kind: 'whole',
    op: 'between',
    a: 1,
    b: 10,
    allowBlank: false,
  };

  it('returns false when there is no validation', () => {
    expect(detectValidationViolation({ kind: 'number', value: 5 }, undefined)).toBe(false);
  });

  it('returns true when the value is out of range', () => {
    expect(detectValidationViolation({ kind: 'number', value: 99 }, wholeRange)).toBe(true);
  });

  it('returns false when the value satisfies the rule', () => {
    expect(detectValidationViolation({ kind: 'number', value: 5 }, wholeRange)).toBe(false);
  });

  it('skips error-kind values — those surface as error triangles instead', () => {
    expect(detectValidationViolation({ kind: 'error', code: 7, text: '#DIV/0!' }, wholeRange)).toBe(
      false,
    );
  });
});

describe('paintErrorTriangle / paintValidationTriangle', () => {
  it('paintErrorTriangle uses the supplied color and emits a filled triangle path', () => {
    const { ctx, calls, fills } = makeStubCtx();
    const rect = paintErrorTriangle(ctx, { x: 10, y: 20, w: 80, h: 22 }, ERROR_TRIANGLE_COLOR);
    expect(fills).toContain(ERROR_TRIANGLE_COLOR);
    expect(calls).toContain('beginPath');
    expect(calls).toContain('moveTo');
    expect(calls).toContain('lineTo');
    expect(calls).toContain('fill');
    // Returned hit-rect is anchored at the cell's top-left.
    expect(rect.x).toBe(10);
    expect(rect.y).toBe(20);
    expect(rect.w).toBeGreaterThan(0);
    expect(rect.h).toBeGreaterThan(0);
  });

  it('paintValidationTriangle defaults to the red sentinel and is otherwise identical', () => {
    const { ctx, fills } = makeStubCtx();
    paintValidationTriangle(ctx, { x: 0, y: 0, w: 50, h: 18 });
    expect(fills).toContain('#d24545');
    expect(fills).toContain(VALIDATION_TRIANGLE_COLOR);
  });
});

/* Black-box exercise of the triangle paint decision. We replicate the
 * grid's per-cell predicate (no canvas needed) so we can assert that the
 * ignoredErrors set wins over an active error. */
function shouldPaintTriangle(
  value: CellValue,
  ignored: ReadonlySet<string>,
  addr = { sheet: 0, row: 0, col: 0 },
): 'error' | null {
  if (ignored.has(addrKey(addr))) return null;
  return detectErrorKind(value) ? 'error' : null;
}

describe('ignoredErrors suppression', () => {
  it('emits a paint when the cell is an error and not ignored', () => {
    const out = shouldPaintTriangle({ kind: 'error', code: 4, text: '#REF!' }, new Set());
    expect(out).toBe('error');
  });

  it('skips the paint once ignoreError is recorded for the addr', () => {
    const store = createSpreadsheetStore();
    mutators.ignoreError(store, { sheet: 0, row: 0, col: 0 });
    const out = shouldPaintTriangle(
      { kind: 'error', code: 4, text: '#REF!' },
      store.getState().errorIndicators.ignoredErrors,
    );
    expect(out).toBeNull();
  });

  it('clearIgnoredErrors restores the paint', () => {
    const store = createSpreadsheetStore();
    mutators.ignoreError(store, { sheet: 0, row: 0, col: 0 });
    mutators.clearIgnoredErrors(store);
    const out = shouldPaintTriangle(
      { kind: 'error', code: 4, text: '#REF!' },
      store.getState().errorIndicators.ignoredErrors,
    );
    expect(out).toBe('error');
  });
});

/* End-to-end: feed a fake validation-failing value through the same pair
 * the renderer uses (`detectValidationViolation` + `paintValidationTriangle`)
 * and confirm we end up with one red triangle hit recorded. This mirrors
 * what `grid.ts` does inside `paintCells`. */
describe('validation failure surface', () => {
  it('a value outside the whole-number range triggers paintValidationTriangle', () => {
    const validation: CellValidation = {
      kind: 'whole',
      op: 'between',
      a: 1,
      b: 10,
      allowBlank: false,
    };
    const value: CellValue = { kind: 'number', value: 99 };
    expect(detectValidationViolation(value, validation)).toBe(true);

    const { ctx, fills } = makeStubCtx();
    const hits: ErrorTriangleHit[] = [];
    if (detectValidationViolation(value, validation)) {
      const rect = paintValidationTriangle(ctx, { x: 0, y: 0, w: 60, h: 18 });
      hits.push({ rect, addr: { sheet: 0, row: 1, col: 1 }, kind: 'validation' });
    }
    expect(hits).toHaveLength(1);
    expect(hits[0]?.kind).toBe('validation');
    expect(fills).toContain(VALIDATION_TRIANGLE_COLOR);
  });
});
