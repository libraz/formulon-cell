import { StatusCode } from '@libraz/formulon';
import { mount, type SpreadsheetInstance, Workbook } from '@libraz/formulon-cell';
import { mountToolbar } from '@libraz/formulon-cell/toolbar';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { expectFailure, expectFailureAsync, makeHost } from './helpers.js';

const LIFECYCLE_MEMBERS = new Set<PropertyKey>(['disposed', 'dispose', Symbol.dispose]);

/** Every own and prototype member of `target` other than the lifecycle trio and the constructor. */
const membersOf = (target: object): PropertyKey[] => {
  const keys = new Set<PropertyKey>();
  for (let o: object | null = target; o && o !== Object.prototype; o = Object.getPrototypeOf(o)) {
    for (const key of Reflect.ownKeys(o)) {
      if (key !== 'constructor' && !LIFECYCLE_MEMBERS.has(key)) keys.add(key);
    }
  }
  return [...keys];
};

/** Touches `key` on `target` the way a consumer would: reads it, and calls it when it is a method. */
const touch = (target: Record<PropertyKey, unknown>, key: PropertyKey): void => {
  const member = target[key];
  if (typeof member === 'function') (member as () => unknown).call(target);
};

const expectEveryMemberDisposed = (target: object): void => {
  for (const key of membersOf(target)) {
    expectFailure(() => touch(target as Record<PropertyKey, unknown>, key), 'disposed');
  }
};

const live: Array<{ dispose(): void }> = [];

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
  vi.unstubAllGlobals();
});

describe('lifecycle: dispose', () => {
  it('Workbook: dispose is idempotent, never throws, and disables every other member', async () => {
    const wb = await Workbook.createDefault();
    expect(typeof wb[Symbol.dispose]).toBe('function');
    wb.dispose();
    wb.dispose();
    wb[Symbol.dispose]();
    expect(wb.disposed).toBe(true);
    const error = expectFailure(() => wb.sheet(0), 'disposed');
    expect(error.code).toBe(StatusCode.BindingInvalidHandle);
    expectEveryMemberDisposed(wb);
  });

  it('SpreadsheetInstance: dispose is idempotent and disables every other member', async () => {
    const instance = await mount(makeHost());
    expect(typeof instance[Symbol.dispose]).toBe('function');
    instance.dispose();
    instance.dispose();
    instance[Symbol.dispose]();
    expect(instance.disposed).toBe(true);
    expectEveryMemberDisposed(instance);
  });

  it('Toolbar: dispose is idempotent and disables every other member', async () => {
    const instance = await mount(makeHost());
    live.push(instance);
    const toolbar = await mountToolbar(makeHost(), instance, {});
    expect(typeof toolbar[Symbol.dispose]).toBe('function');
    toolbar.dispose();
    toolbar.dispose();
    expect(toolbar.disposed).toBe(true);
    expectEveryMemberDisposed(toolbar);
  });

  it('an instance disposes a workbook it created and leaves a passed-in one to its caller', async () => {
    const owned = await mount(makeHost());
    const created = owned.workbook;
    owned.dispose();
    expect(created.disposed).toBe(true);

    const wb = await Workbook.createDefault();
    live.push(wb);
    const borrowed = await mount(makeHost(), { workbook: wb });
    borrowed.dispose();
    expect(wb.disposed).toBe(false);
    expect(wb.listSheets()).toHaveLength(1);
  });
});

describe('lifecycle: events', () => {
  it('on() returns an idempotent disposer on every observable object', async () => {
    const instance = await mount(makeHost(), { locale: 'en' });
    live.push(instance);
    const subscriptions: Array<() => void> = [
      instance.on('update', () => {}),
      instance.workbook.on('change', () => {}),
      instance.selection.on('change', () => {}),
      instance.history.on('change', () => {}),
      instance.i18n.on('change', () => {}),
      instance.cells.on('change', () => {}),
      instance.formulaHints.on('change', () => {}),
    ];
    for (const off of subscriptions) {
      expect(typeof off).toBe('function');
      off();
      off();
    }
  });

  it('a disposed listener no longer runs', async () => {
    const wb = await Workbook.createDefault();
    live.push(wb);
    const listener = vi.fn();
    const off = wb.on('change', listener);
    off();
    wb.sheet(0).setNumber('A1', 1);
    expect(listener).not.toHaveBeenCalled();
  });

  it('a throwing listener reaches reportError and the other listeners still run', async () => {
    const reportError = vi.fn();
    vi.stubGlobal('reportError', reportError);
    const wb = await Workbook.createDefault();
    live.push(wb);
    const boom = new Error('listener failed');
    const after = vi.fn();
    wb.on('change', () => {
      throw boom;
    });
    wb.on('change', after);
    wb.sheet(0).setNumber('A1', 1);
    expect(after).toHaveBeenCalledTimes(1);
    expect(reportError).toHaveBeenCalledWith(boom);
  });

  it('without reportError a throwing listener is rethrown on a microtask', async () => {
    vi.stubGlobal('reportError', undefined);
    const wb = await Workbook.createDefault();
    live.push(wb);
    const boom = new Error('listener failed');
    const queued: Array<() => void> = [];
    const queueMicrotask = vi
      .spyOn(globalThis, 'queueMicrotask')
      .mockImplementation((task) => queued.push(task));
    const after = vi.fn();
    wb.on('change', () => {
      throw boom;
    });
    wb.on('change', after);
    wb.sheet(0).setNumber('A1', 1);
    queueMicrotask.mockRestore();
    expect(after).toHaveBeenCalledTimes(1);
    expect(queued).toHaveLength(1);
    expect(() => queued[0]?.()).toThrow(boom);
  });
});

describe('lifecycle: async factories', () => {
  it('an already aborted signal rejects mount with its reason before touching the host', async () => {
    const host = makeHost();
    const controller = new AbortController();
    const reason = new Error('cancelled');
    controller.abort(reason);
    await expect(mount(host, { signal: controller.signal })).rejects.toBe(reason);
    expect(host.childElementCount).toBe(0);
  });

  it('an already aborted signal rejects createDefault with its reason', async () => {
    const controller = new AbortController();
    controller.abort('stop');
    await expect(Workbook.createDefault({ signal: controller.signal })).rejects.toBe('stop');
  });

  it('a second mount of a mounted workbook throws workbook-in-use', async () => {
    const wb = await Workbook.createDefault();
    live.push(wb);
    const first: SpreadsheetInstance = await mount(makeHost(), { workbook: wb });
    live.push(first);
    await expectFailureAsync(mount(makeHost(), { workbook: wb }), 'workbook-in-use');
    first.dispose();
    const second = await mount(makeHost(), { workbook: wb });
    live.push(second);
    expect(second.workbook).toBe(wb);
  });

  it('workbook and seed together are rejected with TypeError', async () => {
    const wb = await Workbook.createDefault();
    live.push(wb);
    await expect(
      mount(makeHost(), { workbook: wb, seed: () => {} } as never),
    ).rejects.toBeInstanceOf(TypeError);
  });
});
