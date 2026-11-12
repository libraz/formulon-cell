import { describe, expect, it } from 'vitest';
import { PivotAxis, PivotFilterKind, type PivotFilterSpec } from '../../../src/engine/types.js';
import { retainedPivotFilters } from '../../../src/interact/workbook-objects-pivot-editor.js';

const spec = (axis: PivotAxis, fieldName: string): PivotFilterSpec => ({
  axis,
  fieldName,
  kind: PivotFilterKind.LabelContains,
  value: 'x',
});

const placed = (...entries: [string, PivotAxis | null][]) =>
  entries.map(([sourceName, axis]) => ({ editor: { model: { sourceName } }, axis }));

const READ = [spec(PivotAxis.Row, 'Region'), spec(PivotAxis.Col, 'Product')];

describe('retainedPivotFilters', () => {
  it('returns null when every filter stays on a row or column field', () => {
    const next = placed(
      ['Region', PivotAxis.Row],
      ['Product', PivotAxis.Col],
      ['Amt', PivotAxis.Value],
    );
    expect(retainedPivotFilters(READ, next)).toBeNull();
    expect(retainedPivotFilters([], next)).toBeNull();
  });

  it('keeps a filter when its field swaps between rows and columns', () => {
    const next = placed(['Region', PivotAxis.Col], ['Product', PivotAxis.Row]);
    expect(retainedPivotFilters(READ, next)).toBeNull();
  });

  it('drops the filter of a field moved to Filters', () => {
    const next = placed(['Region', PivotAxis.Page], ['Product', PivotAxis.Col]);
    expect(retainedPivotFilters(READ, next)).toEqual([READ[1]]);
  });

  it('drops the filter of a field moved to Values', () => {
    const next = placed(['Region', PivotAxis.Row], ['Product', PivotAxis.Value]);
    expect(retainedPivotFilters(READ, next)).toEqual([READ[0]]);
  });

  it('drops the filter of a field taken out of the report', () => {
    const next = placed(['Region', null], ['Product', PivotAxis.Col]);
    expect(retainedPivotFilters(READ, next)).toEqual([READ[1]]);
  });

  it('drops a filter read on a report filter field', () => {
    const next = placed(
      ['Region', PivotAxis.Row],
      ['Product', PivotAxis.Col],
      ['Amt', PivotAxis.Page],
    );
    expect(retainedPivotFilters([...READ, spec(PivotAxis.Page, 'Amt')], next)).toEqual(READ);
  });
});
