import type { CellValue } from '../../../src/engine/types.js';
import type { ConditionalRule, State } from '../../../src/store/store.js';

const seedCell = (state: State, row: number, col: number, value: CellValue): State => {
  const cells = new Map(state.data.cells);
  cells.set(`0:${row}:${col}`, { value, formula: null });
  return { ...state, data: { ...state.data, cells } };
};

const seedNumber = (state: State, row: number, col: number, value: number): State => {
  const cells = new Map(state.data.cells);
  cells.set(`0:${row}:${col}`, { value: { kind: 'number', value }, formula: null });
  return { ...state, data: { ...state.data, cells } };
};

const dateSerial = (year: number, month: number, day: number): number =>
  Date.UTC(year, month - 1, day) / 86_400_000 + 25569;

const cellValueRule = (range: ConditionalRule['range']): ConditionalRule => ({
  kind: 'cell-value',
  range,
  op: '>',
  a: 5,
  apply: { fill: '#ff0000' },
});

const formulaRule = (
  range: ConditionalRule['range'],
  formula: string,
  fill: string,
): ConditionalRule => ({
  kind: 'formula',
  range,
  formula,
  apply: { fill },
});

const withConditionalRules = (state: State, rules: ConditionalRule[]): State => ({
  ...state,
  conditional: { ...state.conditional, rules },
});

export { cellValueRule, dateSerial, formulaRule, seedCell, seedNumber, withConditionalRules };
