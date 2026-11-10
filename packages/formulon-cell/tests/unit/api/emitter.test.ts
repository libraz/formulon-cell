import { FormulonError, StatusCode } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { spreadsheetFailure, spreadsheetFailureOf } from '../../../src/api/errors.js';
import { Emitter } from '../../../src/api/events.js';
import { Lifecycle } from '../../../src/api/lifecycle.js';

interface Events {
  change: { n: number };
  other: string;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('Emitter', () => {
  it('delivers to listeners in subscription order and stops after the disposer runs', () => {
    const emitter = new Emitter<Events>();
    const seen: string[] = [];
    const offA = emitter.on('change', (e) => seen.push(`a${e.n}`));
    emitter.on('change', (e) => seen.push(`b${e.n}`));
    emitter.emit('change', { n: 1 });
    offA();
    offA();
    emitter.emit('change', { n: 2 });
    expect(seen).toEqual(['a1', 'b1', 'b2']);
  });

  it('keeps two subscriptions of the same function independent', () => {
    const emitter = new Emitter<Events>();
    const listener = vi.fn();
    const off = emitter.on('other', listener);
    emitter.on('other', listener);
    off();
    emitter.emit('other', 'x');
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it('does not call a listener added or removed during an emit for that emit', () => {
    const emitter = new Emitter<Events>();
    const late = vi.fn();
    const removed = vi.fn();
    let offRemoved = (): void => {};
    emitter.on('change', () => {
      emitter.on('change', late);
      offRemoved();
    });
    offRemoved = emitter.on('change', removed);
    emitter.emit('change', { n: 1 });
    expect(late).not.toHaveBeenCalled();
    expect(removed).not.toHaveBeenCalled();
  });

  it('reports a throwing listener through reportError and still runs the rest', () => {
    const reportError = vi.fn();
    vi.stubGlobal('reportError', reportError);
    const emitter = new Emitter<Events>();
    const boom = new Error('boom');
    const after = vi.fn();
    emitter.on('change', () => {
      throw boom;
    });
    emitter.on('change', after);
    emitter.emit('change', { n: 1 });
    expect(reportError).toHaveBeenCalledWith(boom);
    expect(after).toHaveBeenCalledTimes(1);
  });

  it('rethrows on a microtask when reportError is not defined', () => {
    vi.stubGlobal('reportError', undefined);
    const queued: Array<() => void> = [];
    const spy = vi.spyOn(globalThis, 'queueMicrotask').mockImplementation((task) => {
      queued.push(task);
    });
    const emitter = new Emitter<Events>();
    const boom = new Error('boom');
    emitter.on('change', () => {
      throw boom;
    });
    emitter.emit('change', { n: 1 });
    spy.mockRestore();
    expect(queued).toHaveLength(1);
    expect(() => queued[0]?.()).toThrow(boom);
  });
});

describe('spreadsheetFailure', () => {
  it('builds a PreconditionFailed FormulonError whose context leads with the failure pair', () => {
    const cause = new Error('inner');
    const error = spreadsheetFailure('feature-disabled', 'pivot tables are off', {
      context: { feature: 'pivotTable' },
      cause,
    });
    expect(error).toBeInstanceOf(FormulonError);
    expect(error.code).toBe(StatusCode.PreconditionFailed);
    expect(error.context).toBe('failure=feature-disabled feature=pivotTable');
    expect(error.cause).toBe(cause);
    expect(spreadsheetFailureOf(error)).toBe('feature-disabled');
  });

  it('uses BindingInvalidHandle for disposed', () => {
    const error = spreadsheetFailure('disposed', 'gone');
    expect(error.code).toBe(StatusCode.BindingInvalidHandle);
    expect(error.context).toBe('failure=disposed');
    expect('cause' in error).toBe(false);
  });

  it('spreadsheetFailureOf ignores engine errors, unknown kinds and non-errors', () => {
    expect(
      spreadsheetFailureOf(new FormulonError(StatusCode.NotFound, 'x', 'sheet_index=5')),
    ).toBeNull();
    expect(
      spreadsheetFailureOf(new FormulonError(StatusCode.PreconditionFailed, 'x', 'failure=bogus')),
    ).toBeNull();
    expect(
      spreadsheetFailureOf(
        new FormulonError(StatusCode.PreconditionFailed, 'x', 'sheet=0 failure=protected'),
      ),
    ).toBe('protected');
    expect(spreadsheetFailureOf(new Error('failure=protected'))).toBeNull();
    expect(spreadsheetFailureOf(null)).toBeNull();
  });
});

describe('Lifecycle', () => {
  it('runs teardown once, newest first, and guards members afterwards', () => {
    const lifecycle = new Lifecycle('Workbook');
    const order: string[] = [];
    lifecycle.onDispose(() => order.push('first'));
    lifecycle.onDispose(() => order.push('second'));
    expect(() => lifecycle.assertLive()).not.toThrow();
    lifecycle.dispose();
    lifecycle.dispose();
    expect(order).toEqual(['second', 'first']);
    expect(lifecycle.disposed).toBe(true);
    let thrown: unknown;
    try {
      lifecycle.assertLive();
    } catch (e) {
      thrown = e;
    }
    expect(spreadsheetFailureOf(thrown)).toBe('disposed');
  });

  it('reports a failing teardown step instead of throwing and runs the others', () => {
    const reportError = vi.fn();
    vi.stubGlobal('reportError', reportError);
    const lifecycle = new Lifecycle('Toolbar');
    const boom = new Error('boom');
    const other = vi.fn();
    lifecycle.onDispose(other);
    lifecycle.onDispose(() => {
      throw boom;
    });
    expect(() => lifecycle.dispose()).not.toThrow();
    expect(reportError).toHaveBeenCalledWith(boom);
    expect(other).toHaveBeenCalledTimes(1);
  });

  it('runs a step registered after disposal at once', () => {
    const lifecycle = new Lifecycle('Workbook');
    lifecycle.dispose();
    const late = vi.fn();
    lifecycle.onDispose(late);
    expect(late).toHaveBeenCalledTimes(1);
  });
});
