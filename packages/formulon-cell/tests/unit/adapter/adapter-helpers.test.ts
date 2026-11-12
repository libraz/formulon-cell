import { describe, expect, it, vi } from 'vitest';
import {
  forwardRibbonHooks,
  isLiveSpreadsheet,
  optionsFromProps,
  patchFromOptions,
  ribbonHooksKey,
  subscribeToLocale,
  subscribeToSelection,
  subscribeToSpreadsheet,
} from '../../../src/adapter/index.js';
import type { SpreadsheetInstance, SpreadsheetOptions } from '../../../src/index.js';
import type { RibbonHooks } from '../../../src/toolbar/index.js';

describe('adapter option helpers', () => {
  it('keep the set mount options and drop seed and signal from the update patch', () => {
    const signal = new AbortController().signal;
    const props = {
      theme: 'ink',
      locale: undefined,
      seed: [],
      signal,
      children: 'not an option',
    } as unknown as SpreadsheetOptions;
    const options = optionsFromProps(props);
    expect(options).toEqual({ theme: 'ink', seed: [], signal });
    expect(patchFromOptions(options)).toEqual({ theme: 'ink' });
  });
});

describe('adapter ribbon hook helpers', () => {
  it('name the present hooks and forward each call to the hooks held at call time', () => {
    const first = vi.fn(() => 1);
    const second = vi.fn(() => 2);
    let hooks = { insert: { createTable: first, missing: undefined } } as unknown as RibbonHooks;
    expect(ribbonHooksKey(hooks)).toBe('insert.createTable');
    expect(ribbonHooksKey(undefined)).toBe('');
    const forwarded = forwardRibbonHooks(() => hooks) as unknown as {
      insert: { createTable(): unknown };
    };
    hooks = { insert: { createTable: second } } as unknown as RibbonHooks;
    expect(forwarded.insert.createTable()).toBe(2);
    expect(first).not.toHaveBeenCalled();
  });
});

describe('adapter subscribers', () => {
  const fakeSpreadsheet = () => {
    const offs: string[] = [];
    const listen = (name: string) => (_event: string, _fn: () => void) => () => offs.push(name);
    const events: string[] = [];
    const s = {
      disposed: false,
      on: (event: string, fn: () => void) => {
        events.push(event);
        return listen(event)(event, fn);
      },
      selection: { on: listen('selection') },
      i18n: { on: listen('i18n') },
    } as unknown as SpreadsheetInstance;
    return { s, offs, events };
  };

  it('listen to the events each derived value depends on and stop on dispose', () => {
    const all = fakeSpreadsheet();
    subscribeToSpreadsheet(all.s, () => {})();
    expect(all.events).toEqual(['change', 'cell', 'workbook', 'recalc', 'update']);
    expect(all.offs).toEqual([
      'change',
      'cell',
      'workbook',
      'recalc',
      'update',
      'selection',
      'i18n',
    ]);
    const selection = fakeSpreadsheet();
    subscribeToSelection(selection.s, () => {})();
    expect(selection.offs).toEqual(['selection']);
    const locale = fakeSpreadsheet();
    subscribeToLocale(locale.s, () => {})();
    expect(locale.offs).toEqual(['i18n']);
  });

  it('read a value only from a mounted spreadsheet', () => {
    expect(isLiveSpreadsheet(null)).toBe(false);
    expect(isLiveSpreadsheet({ disposed: true } as SpreadsheetInstance)).toBe(false);
    expect(isLiveSpreadsheet({ disposed: false } as SpreadsheetInstance)).toBe(true);
  });
});
