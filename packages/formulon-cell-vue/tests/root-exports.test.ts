import { describe, expect, it } from 'vitest';
import * as vueRoot from '../src/index';

describe('package root', () => {
  it('exports exactly the component, the toolbar and the composables', () => {
    expect(Object.keys(vueRoot).sort()).toEqual([
      'Spreadsheet',
      'SpreadsheetToolbar',
      'useI18n',
      'useSelection',
      'useSpreadsheet',
      'useSpreadsheetEvent',
    ]);
  });
});
