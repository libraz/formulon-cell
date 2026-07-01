import { afterEach, describe, expect, it, vi } from 'vitest';

import type { CellValue } from '../../src/engine/types.js';
import { FormulaRegistry } from '../../src/formula.js';

const numberCell = (value: number): CellValue => ({ kind: 'number', value });

describe('FormulaRegistry', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('stores names case-insensitively and returns a sorted uppercase list', () => {
    const registry = new FormulaRegistry();

    registry.register('tax_rate', () => 0.1);
    registry.register('Discount', () => 5);

    expect(registry.has('TAX_RATE')).toBe(true);
    expect(registry.has('tax_rate')).toBe(true);
    expect(registry.get('discount')?.name).toBe('DISCOUNT');
    expect(registry.list()).toEqual(['DISCOUNT', 'TAX_RATE']);
  });

  it('normalizes host function return values into cell values', () => {
    const registry = new FormulaRegistry();
    registry.register('NUM', () => 12);
    registry.register('TEXT', () => 'ok');
    registry.register('BOOL', () => true);
    registry.register('BLANK', () => null);
    registry.register('CELL', () => ({ kind: 'number', value: 7 }));

    expect(registry.evaluate('num', [])).toEqual({ kind: 'number', value: 12 });
    expect(registry.evaluate('text', [])).toEqual({ kind: 'text', value: 'ok' });
    expect(registry.evaluate('bool', [])).toEqual({ kind: 'bool', value: true });
    expect(registry.evaluate('blank', [])).toEqual({ kind: 'blank' });
    expect(registry.evaluate('cell', [])).toEqual({ kind: 'number', value: 7 });
  });

  it('passes cell arguments through to the registered implementation', () => {
    const registry = new FormulaRegistry();
    const impl = vi.fn((left: CellValue, right: CellValue) =>
      left.kind === 'number' && right.kind === 'number' ? left.value + right.value : null,
    );
    registry.register('ADD_TWO', impl);

    expect(registry.evaluate('add_two', [numberCell(2), numberCell(3)])).toEqual({
      kind: 'number',
      value: 5,
    });
    expect(impl).toHaveBeenCalledWith(numberCell(2), numberCell(3));
  });

  it('lets a disposer remove only the registration it created', () => {
    const registry = new FormulaRegistry();
    const disposeFirst = registry.register('RATE', () => 1);
    registry.register('RATE', () => 2);

    disposeFirst();

    expect(registry.evaluate('RATE', [])).toEqual({ kind: 'number', value: 2 });
  });

  it('notifies subscribers for registry changes and supports unsubscribe', () => {
    const registry = new FormulaRegistry();
    const listener = vi.fn();
    const unsubscribe = registry.subscribe(listener);

    registry.register('ONE', () => 1);
    expect(listener).toHaveBeenCalledTimes(1);

    registry.unregister('one');
    expect(listener).toHaveBeenCalledTimes(2);

    unsubscribe();
    registry.register('TWO', () => 2);
    expect(listener).toHaveBeenCalledTimes(2);
  });

  it('rejects invalid names and unknown evaluations with clear errors', () => {
    const registry = new FormulaRegistry();

    expect(() => registry.register('1INVALID', () => 1)).toThrow(
      'formulon-cell: invalid function name "1INVALID"',
    );
    expect(() => registry.evaluate('MISSING', [])).toThrow(
      'formulon-cell: unknown custom function "MISSING"',
    );
  });

  it('continues notifying other subscribers when one subscriber throws', () => {
    const registry = new FormulaRegistry();
    const error = new Error('listener failed');
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const good = vi.fn();
    registry.subscribe(() => {
      throw error;
    });
    registry.subscribe(good);

    registry.register('SAFE', () => 1);

    expect(good).toHaveBeenCalledTimes(1);
    expect(consoleError).toHaveBeenCalledWith(
      'formulon-cell: formula-registry listener threw',
      error,
    );
  });
});
