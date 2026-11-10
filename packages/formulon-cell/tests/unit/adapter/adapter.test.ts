import '../../api/setup-dom.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { propsPatch, reconcileProps, sameProp } from '../../../src/adapter/index.js';
import { mount, type SpreadsheetInstance } from '../../../src/index.js';

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const i of live.splice(0)) i.dispose();
});

describe('sameProp', () => {
  it('treats inline literals as equal', () => {
    expect(sameProp({ a: 1, b: [1, 2] }, { a: 1, b: [1, 2] })).toBe(true);
    expect(sameProp({ a: { b: { c: 'x' } } }, { a: { b: { c: 'x' } } })).toBe(true);
    expect(sameProp(1, 1)).toBe(true);
    expect(sameProp(Number.NaN, Number.NaN)).toBe(true);
    expect(sameProp(null, undefined)).toBe(false);
  });

  it('tells different shapes and values apart', () => {
    expect(sameProp({ a: 1 }, { a: 1, b: 2 })).toBe(false);
    expect(sameProp({ a: 1 }, { b: 1 })).toBe(false);
    expect(sameProp([1, 2], [1, 2, 3])).toBe(false);
    expect(sameProp({ a: { b: 1 } }, { a: { b: 2 } })).toBe(false);
    expect(sameProp([], {})).toBe(false);
  });

  it('compares functions, class instances and collections by identity', () => {
    const fn = () => 1;
    expect(sameProp(fn, fn)).toBe(true);
    expect(
      sameProp(
        () => 1,
        () => 1,
      ),
    ).toBe(false);
    class Box {
      readonly v = 1;
    }
    const box = new Box();
    expect(sameProp(box, box)).toBe(true);
    expect(sameProp(new Box(), new Box())).toBe(false);
    expect(sameProp(new Map(), new Map())).toBe(false);
    expect(sameProp({ m: new Set() }, { m: new Set() })).toBe(false);
    const signal = new AbortController().signal;
    expect(sameProp({ signal }, { signal })).toBe(true);
  });

  it('terminates on cycles', () => {
    const a: Record<string, unknown> = { n: 1 };
    a.self = a;
    const b: Record<string, unknown> = { n: 1 };
    b.self = b;
    expect(sameProp(a, b)).toBe(true);
    const c: Record<string, unknown> = { n: 2 };
    c.self = c;
    expect(sameProp(a, c)).toBe(false);
  });
});

describe('propsPatch', () => {
  it('keeps changed keys and drops unchanged inline literals', () => {
    const patch = propsPatch(
      { theme: 'ink', policy: { readOnly: false } as never },
      { theme: 'paper', policy: { readOnly: false } as never },
    );
    expect(patch).toEqual({ theme: 'paper' });
  });

  it('maps a removed key to null', () => {
    expect(propsPatch({ theme: 'ink', locale: 'en' }, { locale: 'en' })).toEqual({ theme: null });
    expect(propsPatch({ theme: 'ink' }, { theme: undefined })).toEqual({ theme: null });
    expect(propsPatch({}, { theme: undefined })).toEqual({});
  });
});

describe('reconcileProps', () => {
  const mounted = async (): Promise<SpreadsheetInstance> => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const instance = await mount(host);
    live.push(instance);
    return instance;
  };
  // An open modal dialog makes the instance refuse a locale change.
  const refuseLocale = (instance: SpreadsheetInstance): void => {
    const modal = document.createElement('div');
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    instance.host.append(modal);
  };

  it('applies each changed key with one update event apiece', async () => {
    const instance = await mounted();
    const events: (readonly string[])[] = [];
    instance.on('update', (e) => events.push(e.keys));
    const applied = reconcileProps(instance, {}, { theme: 'ink', locale: 'ja' });
    expect(applied).toEqual(['theme', 'locale']);
    expect(events).toEqual([['theme'], ['locale']]);
    expect(instance.i18n.locale).toBe('ja');
  });

  it('reports a refused key and still applies the others', async () => {
    const instance = await mounted();
    refuseLocale(instance);
    const onError = vi.fn();
    const applied = reconcileProps(instance, {}, { locale: 'ja', theme: 'ink' }, onError);
    expect(applied).toEqual(['theme']);
    expect(onError).toHaveBeenCalledOnce();
    expect(onError.mock.calls[0]?.[1]).toBe('locale');
  });

  it('rethrows the first refusal after applying the rest when no handler is given', async () => {
    const instance = await mounted();
    refuseLocale(instance);
    const events: (readonly string[])[] = [];
    instance.on('update', (e) => events.push(e.keys));
    expect(() => reconcileProps(instance, {}, { locale: 'ja', theme: 'ink' })).toThrow();
    expect(events).toEqual([['theme']]);
  });
});
