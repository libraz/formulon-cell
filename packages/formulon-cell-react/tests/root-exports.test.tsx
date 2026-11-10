import { describe, expect, it } from 'vitest';
import * as ReactPackage from '../src';

describe('React package root', () => {
  it('exports the component, the toolbar and the hooks, and nothing from the core', () => {
    expect(Object.keys(ReactPackage).sort()).toEqual([
      'Spreadsheet',
      'SpreadsheetToolbar',
      'useI18n',
      'useSelection',
      'useSpreadsheet',
      'useSpreadsheetEvent',
    ]);
  });
});
