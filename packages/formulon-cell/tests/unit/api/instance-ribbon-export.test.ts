import { describe, expectTypeOf, it } from 'vitest';
import type { InstanceRibbon, SpreadsheetInstance } from '../../../src/index.js';
import type { InstanceRibbon as ToolbarEntryInstanceRibbon } from '../../../src/toolbar/index.js';

describe('InstanceRibbon', () => {
  it('is the public type of instance.toolbar', () => {
    expectTypeOf<SpreadsheetInstance['toolbar']>().toEqualTypeOf<InstanceRibbon | null>();
    expectTypeOf<ToolbarEntryInstanceRibbon>().toEqualTypeOf<InstanceRibbon>();
  });
});
