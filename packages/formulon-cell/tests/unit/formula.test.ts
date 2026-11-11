import { afterEach, describe, expect, it, vi } from 'vitest';

import { FormulaRegistry } from '../../src/formula.js';

describe('FormulaRegistry', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('stores names case-insensitively and returns a sorted uppercase list', () => {
    const registry = new FormulaRegistry();

    registry.register('tax_rate');
    registry.register('Discount', { args: ['amount'] });

    expect(registry.has('TAX_RATE')).toBe(true);
    expect(registry.has('tax_rate')).toBe(true);
    expect(registry.get('discount')).toEqual({ name: 'DISCOUNT', meta: { args: ['amount'] } });
    expect(registry.list()).toEqual(['DISCOUNT', 'TAX_RATE']);
  });

  it('lets a disposer remove only the registration it created', () => {
    const registry = new FormulaRegistry();
    const disposeFirst = registry.register('RATE', { description: 'first' });
    registry.register('RATE', { description: 'second' });

    disposeFirst();

    expect(registry.get('RATE')?.meta.description).toBe('second');
  });

  it('notifies subscribers for registry changes and supports unsubscribe', () => {
    const registry = new FormulaRegistry();
    const listener = vi.fn();
    const unsubscribe = registry.subscribe(listener);

    registry.register('ONE');
    expect(listener).toHaveBeenCalledTimes(1);

    registry.unregister('one');
    expect(listener).toHaveBeenCalledTimes(2);

    unsubscribe();
    registry.register('TWO');
    expect(listener).toHaveBeenCalledTimes(2);
  });

  it('rejects invalid names with a clear error', () => {
    const registry = new FormulaRegistry();

    expect(() => registry.register('1INVALID')).toThrow(
      'formulon-cell: invalid function name "1INVALID"',
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

    registry.register('SAFE');

    expect(good).toHaveBeenCalledTimes(1);
    expect(consoleError).toHaveBeenCalledWith(
      'formulon-cell: formula-registry listener threw',
      error,
    );
  });
});
