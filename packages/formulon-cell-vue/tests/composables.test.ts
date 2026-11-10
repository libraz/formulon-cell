import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { type EffectScope, effectScope, nextTick, shallowRef, watch } from 'vue';
import { useI18n, useSelection, useSpreadsheet, useSpreadsheetEvent } from '../src';
import { flush, mountStub } from './test-utils/mount';

const scopes: EffectScope[] = [];
const instances: SpreadsheetInstance[] = [];

const mountInstance = async (
  options: Parameters<typeof mountStub>[0] = {},
): Promise<SpreadsheetInstance> => {
  const inst = await mountStub(options);
  instances.push(inst);
  return inst;
};

/** Runs `fn` in a scope the test stops itself or `afterEach` does. */
const inScope = <T>(fn: () => T): { value: T; scope: EffectScope } => {
  const scope = effectScope();
  scopes.push(scope);
  return { value: scope.run(fn) as T, scope };
};

afterEach(() => {
  for (const scope of scopes.splice(0)) scope.stop();
  for (const inst of instances.splice(0).reverse()) inst.dispose();
  document.body.replaceChildren();
});

describe('useSelection', () => {
  it('is null for a null source and mirrors the live selection after', async () => {
    const inst = await mountInstance();
    const source = shallowRef<SpreadsheetInstance | null>(null);
    const { value: selection } = inScope(() => useSelection(source));
    expect(selection.value).toBeNull();
    source.value = inst;
    await nextTick();
    expect(selection.value?.active).toEqual({ row: 0, col: 0 });
    inst.selection.set('C3');
    expect(selection.value?.active).toEqual({ row: 2, col: 2 });
  });

  it('accepts a plain value, a ref and a getter', async () => {
    const inst = await mountInstance();
    const holder = shallowRef<SpreadsheetInstance | null>(inst);
    const { value } = inScope(() => ({
      plain: useSelection(inst),
      byRef: useSelection(shallowRef<SpreadsheetInstance | null>(inst)),
      byGetter: useSelection(() => holder.value),
    }));
    inst.selection.set('B2');
    for (const selection of Object.values(value)) {
      expect(selection.value?.active).toEqual({ row: 1, col: 1 });
    }
  });

  it('follows a new instance and ignores the old one', async () => {
    const a = await mountInstance();
    const b = await mountInstance();
    const source = shallowRef<SpreadsheetInstance | null>(a);
    const { value: selection } = inScope(() => useSelection(source));
    source.value = b;
    await nextTick();
    a.selection.set('D4');
    expect(selection.value?.active).toEqual({ row: 0, col: 0 });
    b.selection.set('E5');
    expect(selection.value?.active).toEqual({ row: 4, col: 4 });
  });

  it('goes null after the instance is disposed', async () => {
    const inst = await mountInstance();
    const { value: selection } = inScope(() => useSelection(inst));
    expect(selection.value).not.toBeNull();
    inst.dispose();
    const source = shallowRef<SpreadsheetInstance | null>(inst);
    const { value: afterDispose } = inScope(() => useSelection(source));
    expect(afterDispose.value).toBeNull();
  });

  it('stops listening when its scope stops', async () => {
    const inst = await mountInstance();
    const { value: selection, scope } = inScope(() => useSelection(inst));
    scope.stop();
    inst.selection.set('F6');
    expect(selection.value?.active).toEqual({ row: 0, col: 0 });
  });
});

describe('useSpreadsheet', () => {
  it('runs the selector now and again after a selection move and an update', async () => {
    const inst = await mountInstance();
    const selector = vi.fn((s: SpreadsheetInstance) => s.selection.get().active.row);
    const { value } = inScope(() => useSpreadsheet(inst, selector));
    expect(value.value).toBe(0);
    inst.selection.set('A5');
    expect(value.value).toBe(4);
    const calls = selector.mock.calls.length;
    inst.update({ theme: 'ink' });
    expect(selector.mock.calls.length).toBeGreaterThan(calls);
  });

  it('re-runs after a cell change', async () => {
    const inst = await mountInstance();
    const { value } = inScope(() =>
      useSpreadsheet(inst, (s) => s.workbook.sheet(0).getTypedValue(0, 0)),
    );
    expect(value.value).not.toMatchObject({ number: 9 });
    inst.workbook.sheet(0).setNumber(0, 0, 9);
    expect(value.value).toMatchObject({ number: 9 });
  });

  it('is null for a null source and after its scope stops', async () => {
    const inst = await mountInstance();
    const source = shallowRef<SpreadsheetInstance | null>(null);
    const selector = vi.fn((s: SpreadsheetInstance) => s.selection.get().active.row);
    const { value, scope } = inScope(() => useSpreadsheet(source, selector));
    expect(value.value).toBeNull();
    expect(selector).not.toHaveBeenCalled();
    source.value = inst;
    await nextTick();
    expect(value.value).toBe(0);
    scope.stop();
    const calls = selector.mock.calls.length;
    inst.selection.set('A9');
    expect(selector.mock.calls.length).toBe(calls);
  });

  it('notifies watchers even when the selector returns the same object', async () => {
    const inst = await mountInstance();
    const shared = { n: 0 };
    const { value } = inScope(() => useSpreadsheet(inst, () => shared));
    const seen = vi.fn();
    inScope(() => watch(value, seen));
    inst.selection.set('B3');
    await nextTick();
    expect(seen).toHaveBeenCalled();
  });
});

describe('useI18n', () => {
  it('is null before ready, then the active locale and strings', async () => {
    const inst = await mountInstance({ locale: 'en' });
    const source = shallowRef<SpreadsheetInstance | null>(null);
    const { value } = inScope(() => useI18n(source));
    expect(value.value).toBeNull();
    source.value = inst;
    await nextTick();
    expect(value.value?.locale).toBe('en');
    expect(value.value?.strings).toBe(inst.i18n.strings);
  });

  it('updates when the locale changes through update', async () => {
    const inst = await mountInstance({ locale: 'en' });
    const { value } = inScope(() => useI18n(inst));
    inst.update({ locale: 'ja' });
    await flush();
    expect(value.value?.locale).toBe('ja');
    expect(value.value?.strings).toBe(inst.i18n.strings);
  });

  it('follows a new instance, ignores the old one, and stops with its scope', async () => {
    const a = await mountInstance({ locale: 'en' });
    const b = await mountInstance({ locale: 'ja' });
    const source = shallowRef<SpreadsheetInstance | null>(a);
    const { value, scope } = inScope(() => useI18n(source));
    source.value = b;
    await nextTick();
    expect(value.value?.locale).toBe('ja');
    a.update({ locale: 'ko' });
    await flush();
    expect(value.value?.locale).toBe('ja');
    scope.stop();
    b.update({ locale: 'en' });
    await flush();
    expect(value.value?.locale).toBe('ja');
  });
});

describe('useSpreadsheetEvent', () => {
  it('forwards payloads to the handler and follows the source', async () => {
    const a = await mountInstance();
    const b = await mountInstance();
    const source = shallowRef<SpreadsheetInstance | null>(null);
    const handler = vi.fn();
    inScope(() => useSpreadsheetEvent(source, 'update', handler));
    a.update({ theme: 'ink' });
    expect(handler).not.toHaveBeenCalled();
    source.value = a;
    await nextTick();
    a.update({ theme: 'paper' });
    expect(handler).toHaveBeenCalledWith({ keys: ['theme'] });
    source.value = b;
    await nextTick();
    handler.mockClear();
    a.update({ theme: 'ink' });
    expect(handler).not.toHaveBeenCalled();
    b.update({ theme: 'ink' });
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it('accepts a getter and a plain instance', async () => {
    const inst = await mountInstance();
    const byGetter = vi.fn();
    const byValue = vi.fn();
    inScope(() => {
      useSpreadsheetEvent(() => inst, 'update', byGetter);
      useSpreadsheetEvent(inst, 'update', byValue);
    });
    inst.update({ theme: 'ink' });
    expect(byGetter).toHaveBeenCalledTimes(1);
    expect(byValue).toHaveBeenCalledTimes(1);
  });

  it('stops invoking the handler once its scope stops', async () => {
    const inst = await mountInstance();
    const handler = vi.fn();
    const { scope } = inScope(() => useSpreadsheetEvent(inst, 'update', handler));
    scope.stop();
    inst.update({ theme: 'ink' });
    expect(handler).not.toHaveBeenCalled();
  });

  it('does not subscribe to a disposed instance', async () => {
    const inst = await mountInstance();
    inst.dispose();
    const handler = vi.fn();
    expect(() => inScope(() => useSpreadsheetEvent(inst, 'update', handler))).not.toThrow();
    expect(handler).not.toHaveBeenCalled();
  });
});
