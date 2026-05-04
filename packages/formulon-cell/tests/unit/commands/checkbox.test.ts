import { describe, expect, it, vi } from 'vitest';
import { isCheckboxValueChecked, toggleCheckboxValue } from '../../../src/commands/checkbox.js';
import { paintCheckbox } from '../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

describe('toggleCheckboxValue', () => {
  it('blank → checked TRUE', () => {
    expect(toggleCheckboxValue({ kind: 'blank' })).toEqual({ kind: 'bool', value: true });
  });

  it('undefined → checked TRUE', () => {
    expect(toggleCheckboxValue(undefined)).toEqual({ kind: 'bool', value: true });
  });

  it('flips bool', () => {
    expect(toggleCheckboxValue({ kind: 'bool', value: true })).toEqual({
      kind: 'bool',
      value: false,
    });
    expect(toggleCheckboxValue({ kind: 'bool', value: false })).toEqual({
      kind: 'bool',
      value: true,
    });
  });

  it('treats numeric 0 as unchecked → TRUE', () => {
    expect(toggleCheckboxValue({ kind: 'number', value: 0 })).toEqual({
      kind: 'bool',
      value: true,
    });
  });

  it('treats non-zero as checked → FALSE', () => {
    expect(toggleCheckboxValue({ kind: 'number', value: 1 })).toEqual({
      kind: 'bool',
      value: false,
    });
  });
});

describe('isCheckboxValueChecked', () => {
  it('blank/undefined → false', () => {
    expect(isCheckboxValueChecked(undefined)).toBe(false);
    expect(isCheckboxValueChecked({ kind: 'blank' })).toBe(false);
  });

  it('bool value pass-through', () => {
    expect(isCheckboxValueChecked({ kind: 'bool', value: true })).toBe(true);
    expect(isCheckboxValueChecked({ kind: 'bool', value: false })).toBe(false);
  });

  it('non-zero number → true; zero → false', () => {
    expect(isCheckboxValueChecked({ kind: 'number', value: 0 })).toBe(false);
    expect(isCheckboxValueChecked({ kind: 'number', value: -1 })).toBe(true);
  });

  it('text → true when non-empty, false when empty', () => {
    expect(isCheckboxValueChecked({ kind: 'text', value: '' })).toBe(false);
    expect(isCheckboxValueChecked({ kind: 'text', value: 'x' })).toBe(true);
  });
});

describe('paintCheckbox', () => {
  const theme = {
    bg: '#ffffff',
    fg: '#222222',
    accent: '#1f7ae0',
    ruleStrong: '#888888',
    ruleSoft: '#cccccc',
  } as unknown as ResolvedTheme;

  function makeStubCtx() {
    const calls: string[] = [];
    const ctx = {
      save: vi.fn(() => calls.push('save')),
      restore: vi.fn(() => calls.push('restore')),
      beginPath: vi.fn(() => calls.push('beginPath')),
      moveTo: vi.fn(() => calls.push('moveTo')),
      lineTo: vi.fn(() => calls.push('lineTo')),
      stroke: vi.fn(() => calls.push('stroke')),
      strokeRect: vi.fn(() => calls.push('strokeRect')),
      fillRect: vi.fn(() => calls.push('fillRect')),
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 0,
    };
    return { ctx, calls };
  }

  it('checked: paints filled rect + check stroke', () => {
    const { ctx, calls } = makeStubCtx();
    const bounds = { x: 0, y: 0, w: 100, h: 30 };
    paintCheckbox(ctx as unknown as CanvasRenderingContext2D, bounds, true, theme);
    expect(calls).toContain('fillRect');
    expect(calls).toContain('stroke');
  });

  it('unchecked: paints outlined rect (no stroke path)', () => {
    const { ctx, calls } = makeStubCtx();
    const bounds = { x: 0, y: 0, w: 100, h: 30 };
    paintCheckbox(ctx as unknown as CanvasRenderingContext2D, bounds, false, theme);
    expect(calls).toContain('strokeRect');
    expect(calls).not.toContain('moveTo');
  });

  it('returns a hit rect centered in bounds', () => {
    const { ctx } = makeStubCtx();
    const hit = paintCheckbox(
      ctx as unknown as CanvasRenderingContext2D,
      { x: 0, y: 0, w: 100, h: 30 },
      false,
      theme,
    );
    // Glyph is 14x14, centered in 100x30.
    expect(hit.rect.w).toBe(14);
    expect(hit.rect.h).toBe(14);
    expect(hit.rect.x).toBe(43);
    expect(hit.rect.y).toBe(8);
  });
});
