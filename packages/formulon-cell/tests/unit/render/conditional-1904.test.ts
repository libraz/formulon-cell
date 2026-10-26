import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetConditionalCache, evaluateConditional } from '../../../src/render/conditional.js';
import type { ConditionalRule, State } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import {
  dateSerial,
  formulaRule,
  seedNumber,
  withConditionalRules,
} from './conditional-fixtures.js';

const with1904 = (state: State, date1904: boolean): State => ({
  ...state,
  ui: { ...state.ui, date1904 },
});

describe('conditional formatting in a 1904 workbook', () => {
  afterEach(() => {
    _resetConditionalCache();
    vi.useRealTimers();
  });

  it('evaluates TODAY, DATE and the date parts in the 1904 system', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(Date.UTC(2024, 0, 1, 12)));
    const serial = dateSerial(2024, 1, 1) - 1462;
    let s = with1904(createSpreadsheetStore().getState(), true);
    s = seedNumber(s, 0, 0, serial);
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=A1=TODAY()', '#today'),
      formulaRule({ sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, '=A1=DATE(2024,1,1)', '#date'),
      formulaRule(
        { sheet: 0, r0: 0, c0: 3, r1: 0, c1: 3 },
        '=AND(YEAR(A1)=2024,MONTH(A1)=1,DAY(A1)=1)',
        '#parts',
      ),
      formulaRule({ sheet: 0, r0: 0, c0: 4, r1: 0, c1: 4 }, '=INT(NOW())=A1', '#now'),
    ]);
    const overlay = evaluateConditional(s);
    expect(overlay.get('0:0:1')?.fill).toBe('#today');
    expect(overlay.get('0:0:2')?.fill).toBe('#date');
    expect(overlay.get('0:0:3')?.fill).toBe('#parts');
    expect(overlay.get('0:0:4')?.fill).toBe('#now');
  });

  it('does not match the 1900 serial of today', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(Date.UTC(2024, 0, 1, 12)));
    let s = with1904(createSpreadsheetStore().getState(), true);
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 1));
    s = withConditionalRules(s, [
      formulaRule({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, '=A1=TODAY()', '#today'),
    ]);
    expect(evaluateConditional(s).get('0:0:1')?.fill).toBeUndefined();
  });

  it('matches date-occurring rules against 1904 serials and re-evaluates on a system switch', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(Date.UTC(2024, 0, 1, 12)));
    const rule: ConditionalRule = {
      kind: 'date-occurring',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      period: 'today',
      apply: { fill: '#00ff00' },
    };
    let s = with1904(createSpreadsheetStore().getState(), true);
    s = seedNumber(s, 0, 0, dateSerial(2024, 1, 1) - 1462);
    s = withConditionalRules(s, [rule]);
    expect(evaluateConditional(s).get('0:0:0')?.fill).toBe('#00ff00');
    expect(evaluateConditional(with1904(s, false)).get('0:0:0')?.fill).toBeUndefined();
  });
});
