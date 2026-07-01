import { afterEach, describe, expect, it, vi } from 'vitest';

import { CellRegistry, type CellRenderInput } from '../../src/cells.js';

const input: CellRenderInput = {
  addr: { sheet: 0, row: 0, col: 0 },
  value: { kind: 'number', value: 10 },
  formula: null,
  format: undefined,
};

describe('CellRegistry', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('resolves formatters by priority and falls through on null', () => {
    const registry = new CellRegistry();
    registry.registerFormatter({
      id: 'fallback',
      priority: 50,
      match: () => true,
      format: () => 'fallback',
    });
    registry.registerFormatter({
      id: 'pass',
      priority: 10,
      match: () => true,
      format: () => null,
    });
    registry.registerFormatter({
      id: 'winner',
      priority: 20,
      match: ({ value }) => value.kind === 'number',
      format: ({ value }) => (value.kind === 'number' ? `n=${value.value}` : null),
    });

    expect(registry.formatterIds()).toEqual(['pass', 'winner', 'fallback']);
    expect(registry.resolveDisplay(input)).toBe('n=10');
  });

  it('replaces duplicate formatter ids and keeps stale disposers from removing the replacement', () => {
    const registry = new CellRegistry();
    const disposeFirst = registry.registerFormatter({
      id: 'domain',
      match: () => true,
      format: () => 'first',
    });
    const disposeSecond = registry.registerFormatter({
      id: 'domain',
      match: () => true,
      format: () => 'second',
    });

    expect(registry.formatterIds()).toEqual(['domain']);
    expect(registry.resolveDisplay(input)).toBe('second');

    disposeFirst();
    expect(registry.formatterIds()).toEqual(['domain']);
    expect(registry.resolveDisplay(input)).toBe('second');

    disposeSecond();
    expect(registry.formatterIds()).toEqual([]);
    expect(registry.resolveDisplay(input)).toBeNull();
  });

  it('unregisters a formatter by id and reports whether anything changed', () => {
    const registry = new CellRegistry();
    registry.registerFormatter({
      id: 'text',
      match: () => true,
      format: () => 'text',
    });

    expect(registry.unregisterFormatter('missing')).toBe(false);
    expect(registry.unregisterFormatter('text')).toBe(true);
    expect(registry.resolveDisplay(input)).toBeNull();
  });

  it('notifies subscribers for formatter mutations and isolates throwing listeners', () => {
    const registry = new CellRegistry();
    const error = new Error('listener failed');
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const good = vi.fn();
    registry.subscribe(() => {
      throw error;
    });
    const unsubscribe = registry.subscribe(good);

    const dispose = registry.registerFormatter({
      id: 'format',
      match: () => true,
      format: () => 'ok',
    });
    expect(good).toHaveBeenCalledTimes(1);

    unsubscribe();
    dispose();

    expect(good).toHaveBeenCalledTimes(1);
    expect(consoleError).toHaveBeenCalledWith('formulon-cell: cell-registry listener threw', error);
  });

  it('resolves editor entries by priority and removes only the registered entry', () => {
    const registry = new CellRegistry();
    registry.registerEditor({
      id: 'fallback-editor',
      priority: 50,
      match: () => true,
      mount: () => ({ readValue: () => '', focus() {}, detach() {} }),
    });
    const disposeWinner = registry.registerEditor({
      id: 'number-editor',
      priority: 10,
      match: ({ value }) => value.kind === 'number',
      mount: () => ({ readValue: () => '10', focus() {}, detach() {} }),
    });

    expect(registry.resolveEditor(input)?.id).toBe('number-editor');

    disposeWinner();

    expect(registry.resolveEditor(input)?.id).toBe('fallback-editor');
  });
});
