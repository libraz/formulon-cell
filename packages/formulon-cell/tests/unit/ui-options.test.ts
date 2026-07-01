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

  it('applies profile defaults before advanced and friendly feature overrides', () => {
    const resolved = resolveSpreadsheetUiOptions({
      profile: 'minimal',
      advancedFeatures: {
        quickAnalysis: true,
        charts: true,
        validation: false,
      },
      features: {
        quickAnalysis: false,
        validation: true,
        contextMenu: true,
      },
    });

    expect(resolved.profile).toBe('minimal');
    expect(resolved.features.charts).toBe(true);
    expect(resolved.features.quickAnalysis).toBe(false);
    expect(resolved.features.validation).toBe(true);
    expect(resolved.features.contextMenu).toBe(true);
    expect(resolved.features.sheetTabs).toBe(false);
  });

  it('preserves theme and lockTheme host options independently from feature flags', () => {
    const resolved = resolveSpreadsheetUiOptions({
      profile: 'standard',
      theme: 'ink',
      lockTheme: true,
      features: {
        ribbon: false,
      },
    });

    expect(resolved.profile).toBe('standard');
    expect(resolved.theme).toBe('ink');
    expect(resolved.lockTheme).toBe(true);
    expect(resolved.ribbon).toBe(false);
    expect(resolved.features.formatDialog).toBe(false);
  });
});
