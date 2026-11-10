import { describe, expect, it } from 'vitest';
import * as vueRoot from '../src/index';
import { Spreadsheet } from '../src/Spreadsheet';

// The cross-adapter comparison lives in the React package's twin of this file;
// these are the Vue-side conventions it relies on.
describe('Vue adapter surface', () => {
  it('names every emit in camelCase', () => {
    const emits = Object.keys((Spreadsheet as { emits?: Record<string, unknown> }).emits ?? {});
    expect(emits.filter((name) => name.includes('-'))).toEqual([]);
  });

  it('exports the toolbar component from the package root', () => {
    expect(Object.keys(vueRoot)).toContain('SpreadsheetToolbar');
  });
});
