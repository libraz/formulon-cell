import { describe, expect, it } from 'vitest';

import { PivotAxis, PivotFilterKind } from '../../../src/engine/types.js';
import {
  pivotFilterConditionToSpec,
  pivotFilterSpecToCondition,
} from '../../../src/interact/pivot-field-settings.js';

describe('filter condition data field conversion', () => {
  it('writes dataFieldIndex only above the first data field', () => {
    const base = { kind: PivotFilterKind.ValueGreaterThan, value: '5' };
    expect(pivotFilterConditionToSpec('Sales', { ...base, dataField: 2 }, PivotAxis.Row)).toEqual({
      axis: PivotAxis.Row,
      fieldName: 'Sales',
      kind: PivotFilterKind.ValueGreaterThan,
      value: 5,
      dataFieldIndex: 2,
    });
    expect(pivotFilterConditionToSpec('Sales', { ...base, dataField: 0 })).not.toHaveProperty(
      'dataFieldIndex',
    );
  });

  it('reads dataFieldIndex back into the condition', () => {
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Row,
        fieldName: 'Sales',
        kind: PivotFilterKind.ValueTopSum,
        value: 3,
        dataFieldIndex: 1,
      }),
    ).toEqual({ kind: PivotFilterKind.ValueTopSum, value: '3', dataField: 1 });
  });
});
