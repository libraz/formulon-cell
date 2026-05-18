import { describe, expect, it } from 'vitest';
import { resolveSpreadsheetUiOptions } from '../../src/extensions/ui-options.js';

describe('resolveSpreadsheetUiOptions', () => {
  it('defaults to the Excel 365 profile with full chrome enabled', () => {
    const resolved = resolveSpreadsheetUiOptions();

    expect(resolved.profile).toBe('excel365');
    expect(resolved.theme).toBe('paper');
    expect(resolved.ribbon).toBe(true);
    expect(resolved.print).toBe(true);
    expect(resolved.features).toEqual({});
  });

  it('maps user-facing switches to internal feature flags', () => {
    const resolved = resolveSpreadsheetUiOptions({
      features: {
        ribbon: false,
        print: false,
        pivotTable: false,
        slicer: true,
      },
    });

    expect(resolved.ribbon).toBe(false);
    expect(resolved.print).toBe(false);
    expect(resolved.features.pageSetup).toBe(false);
    expect(resolved.features.pivotTableDialog).toBe(false);
    expect(resolved.features.slicer).toBe(true);
  });

  it('lets explicit page setup override print-derived defaults', () => {
    const resolved = resolveSpreadsheetUiOptions({
      features: {
        print: false,
        pageSetup: true,
      },
    });

    expect(resolved.print).toBe(false);
    expect(resolved.features.pageSetup).toBe(true);
  });
});
