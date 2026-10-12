import { afterEach, describe, expect, it } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { _resetConditionalCache, evaluateConditional } from '../../src/render/conditional.js';
import type { ConditionalRule, State } from '../../src/store/store.js';
import { createSpreadsheetStore } from '../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 } as const;
const profileIds = [
  'mac-en_US',
  'mac-ja_JP',
  'mac-de_DE',
  'mac-fr_FR',
  'mac-zh_CN',
  'mac-ko_KR',
  'mac-th_TH',
  'windows-en_US',
  'windows-ja_JP',
  'windows-de_DE',
  'windows-fr_FR',
  'windows-zh_CN',
  'windows-ko_KR',
  'windows-th_TH',
] as const;

const formulaRule = (formula: string): ConditionalRule => ({
  kind: 'formula',
  range,
  formula,
  apply: { fill: '#formula' },
});

const cellValueRule = (
  ruleRange: ConditionalRule['range'],
  op: Extract<ConditionalRule, { kind: 'cell-value' }>['op'],
  a: number | string,
  b?: number | string,
): ConditionalRule => ({
  kind: 'cell-value',
  range: ruleRange,
  op,
  a,
  ...(b === undefined ? {} : { b }),
  apply: { fill: '#cell-value' },
});

const stateWithCells = (cells: State['data']['cells'], formula: string): State => {
  const store = createSpreadsheetStore();
  store.setState((state) => ({
    ...state,
    data: { ...state.data, cells },
    conditional: { rules: [formulaRule(formula)] },
  }));
  return store.getState();
};

describe('integration: native conditional-format formula evaluation', () => {
  afterEach(() => {
    _resetConditionalCache();
  });

  it.each(profileIds)('%s evaluates relative references and functions', async (profile) => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.capabilities.conditionalFormulaEvaluation).toBe(true);
      expect(workbook.setSpreadsheetProfileId(profile)).toBe(true);
      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
      workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 0);
      workbook.setNumber({ sheet: 0, row: 0, col: 2 }, 1);
      const cells = new Map([
        ['0:0:0', { value: { kind: 'number' as const, value: 2 }, formula: null }],
        ['0:0:1', { value: { kind: 'number' as const, value: 0 }, formula: null }],
        ['0:0:2', { value: { kind: 'number' as const, value: 1 }, formula: null }],
      ]);
      const state = stateWithCells(cells, '=AND(SUM(A1,1)>2,ISNUMBER(A1))');
      const overlay = evaluateConditional(state, workbook, profile);

      expect(overlay.get('0:0:0')?.fill).toBe('#formula');
      expect(overlay.get('0:0:1')).toBeUndefined();
    } finally {
      workbook.dispose();
    }
  });

  it('anchors relative references at the top-left of the applied range', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.setSpreadsheetProfileId('mac-en_US')).toBe(true);
      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
      workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 3);
      workbook.setNumber({ sheet: 0, row: 0, col: 2 }, 5);
      const cells = new Map([
        ['0:0:0', { value: { kind: 'number' as const, value: 10 }, formula: null }],
        ['0:0:1', { value: { kind: 'number' as const, value: 3 }, formula: null }],
        ['0:0:2', { value: { kind: 'number' as const, value: 5 }, formula: null }],
      ]);
      const state = stateWithCells(cells, '=A1>$C$1');
      const overlay = evaluateConditional(state, workbook, 'en');

      expect(overlay.get('0:0:0')?.fill).toBe('#formula');
      expect(overlay.get('0:0:1')).toBeUndefined();
    } finally {
      workbook.dispose();
    }
  });

  it.each(profileIds)('%s uses profile collation for text cell-value rules', async (profile) => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.setSpreadsheetProfileId(profile)).toBe(true);
      workbook.setText({ sheet: 0, row: 0, col: 0 }, 'あ');
      const state = stateWithCells(
        new Map([['0:0:0', { value: { kind: 'text' as const, value: 'あ' }, formula: null }]]),
        '=TRUE',
      );
      const textState: State = {
        ...state,
        conditional: {
          rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, '=', 'ア')],
        },
      };
      const overlay = evaluateConditional(textState, workbook, profile);
      const expected = profile.startsWith('mac-') && profile !== 'mac-zh_CN';
      expect(overlay.has('0:0:0')).toBe(expected);
    } finally {
      workbook.dispose();
    }
  });

  it('supports formula-valued boundaries and between operators', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.setSpreadsheetProfileId('mac-en_US')).toBe(true);
      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
      workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 15);
      workbook.setNumber({ sheet: 0, row: 0, col: 2 }, 12);
      const store = createSpreadsheetStore();
      store.setState((state) => ({
        ...state,
        data: {
          ...state.data,
          cells: new Map([
            ['0:0:0', { value: { kind: 'number' as const, value: 5 }, formula: null }],
            ['0:0:1', { value: { kind: 'number' as const, value: 15 }, formula: null }],
            ['0:0:2', { value: { kind: 'number' as const, value: 12 }, formula: null }],
          ]),
        },
        conditional: {
          rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, 'between', 10, 20)],
        },
      }));
      const between = evaluateConditional(store.getState(), workbook, 'en');
      expect(between.has('0:0:0')).toBe(false);
      expect(between.has('0:0:1')).toBe(true);

      store.setState((state) => ({
        ...state,
        conditional: {
          rules: [cellValueRule({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, '>', '=C1')],
        },
      }));
      expect(evaluateConditional(store.getState(), workbook, 'en').has('0:0:0')).toBe(false);
    } finally {
      workbook.dispose();
    }
  });

  it('can bound native predicate work to the painted viewport', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.setSpreadsheetProfileId('mac-en_US')).toBe(true);
      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
      workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 1);
      const state = stateWithCells(
        new Map([
          ['0:0:0', { value: { kind: 'number' as const, value: 1 }, formula: null }],
          ['0:0:1', { value: { kind: 'number' as const, value: 1 }, formula: null }],
        ]),
        '=TRUE',
      );
      const visible = evaluateConditional(state, workbook, 'en', {
        r0: 0,
        c0: 1,
        r1: 0,
        c1: 1,
      });
      expect(visible.has('0:0:0')).toBe(false);
      expect(visible.get('0:0:1')?.fill).toBe('#formula');
    } finally {
      workbook.dispose();
    }
  });

  it('does not reuse a native result after the spreadsheet profile changes', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      const state = stateWithCells(new Map(), '="あ"="ア"');
      expect(workbook.setSpreadsheetProfileId('mac-ja_JP')).toBe(true);
      const ja = evaluateConditional(state, workbook, 'ja');
      expect(ja.get('0:0:0')?.fill).toBe('#formula');

      expect(workbook.setSpreadsheetProfileId('mac-zh_CN')).toBe(true);
      const zh = evaluateConditional(state, workbook, 'zh');
      expect(zh).not.toBe(ja);
      expect(zh.get('0:0:0')).toBeUndefined();
    } finally {
      workbook.dispose();
    }
  });

  it('treats native errors and failed status as no match', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.setSpreadsheetProfileId('mac-en_US')).toBe(true);
      const errorState = stateWithCells(new Map(), '=1/0');
      expect(evaluateConditional(errorState, workbook, 'en').size).toBe(0);

      const failedEngine = {
        capabilities: {
          conditionalFormulaEvaluation: true,
          spreadsheetProfile: false,
        },
        evaluateConditionalFormula: () => ({
          status: { ok: false },
          value: { kind: 1, number: 1, boolean: 0, text: '', errorCode: 0 },
        }),
      } as unknown as WorkbookHandle;
      const fallbackState = stateWithCells(
        new Map([['0:0:0', { value: { kind: 'number' as const, value: 2 }, formula: null }]]),
        '=A1>1',
      );
      expect(evaluateConditional(fallbackState, failedEngine, 'en').size).toBe(0);
    } finally {
      workbook.dispose();
    }
  });

  it('keeps the local predicate path for unsupported or stub engines', () => {
    const state = stateWithCells(
      new Map([['0:0:0', { value: { kind: 'number' as const, value: 2 }, formula: null }]]),
      '=A1>1',
    );
    expect(evaluateConditional(state).get('0:0:0')?.fill).toBe('#formula');
  });
});
