import { describe, expect, it } from 'vitest';

import { SHEET_TAB_COLOR_ACTIONS } from '../../src/index.js';
import {
  SHEET_TAB_COLOR_CHOICES,
  sheetTabColorActionForColor,
  sheetTabColorByAction,
} from '../../src/sheet-tab-colors.js';

describe('sheet tab color shared choices', () => {
  it('keeps ribbon actions, wrapper actions, and color lookup in one palette', () => {
    expect(SHEET_TAB_COLOR_CHOICES.map((choice) => choice.action)).toEqual([
      'tab-color-none',
      'tab-color-red',
      'tab-color-orange',
      'tab-color-yellow',
      'tab-color-green',
      'tab-color-blue',
      'tab-color-purple',
      'tab-color-gray',
    ]);
    expect(sheetTabColorByAction('tab-color-blue')).toBe('#4472c4');
    expect(sheetTabColorByAction('tab-color-none')).toBeNull();
    expect(sheetTabColorByAction('unknown')).toBeUndefined();
    expect(sheetTabColorActionForColor('#4472C4')).toBe('tab-color-blue');
    expect(sheetTabColorActionForColor(undefined)).toBe('tab-color-none');
  });

  it('derives the public wrapper color actions from the same palette', () => {
    expect(SHEET_TAB_COLOR_ACTIONS).toEqual(
      SHEET_TAB_COLOR_CHOICES.filter((choice) => choice.color !== null).map((choice) => ({
        action: choice.wrapperAction,
        color: choice.color,
      })),
    );
  });
});
