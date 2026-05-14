import { afterEach, describe, expect, it } from 'vitest';

import { presets } from '../../src/extensions/presets.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: assert that each documented preset materialises into a
 * specific set of resolved features on the live SpreadsheetInstance.
 * Mirrors README claims so a refactor that drops a default surfaces here.
 */
describe('integration: preset coverage matrix', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => sheet?.dispose());

  it('presets.full() activates the full chrome surface', async () => {
    sheet = await mountStubSheet();
    const { features } = sheet.instance;
    for (const id of [
      'statusBar',
      'viewToolbar',
      'workbookObjects',
      'clipboard',
      'pasteSpecial',
      'quickAnalysis',
      'charts',
      'pivotTableDialog',
      'contextMenu',
      'findReplace',
      'validation',
    ]) {
      expect(features[id], `presets.full() should activate ${id}`).toBeTruthy();
    }
  });

  it('presets.standard() drops authoring dialogs but keeps day-to-day chrome', async () => {
    sheet = await mountStubSheet({ features: presets.standard() });
    const { features } = sheet.instance;
    // kept
    expect(features.clipboard).toBeTruthy();
    expect(features.contextMenu).toBeTruthy();
    expect(features.findReplace).toBeTruthy();
    expect(features.viewToolbar).toBeTruthy();
    expect(features.charts).toBeTruthy();
    // dropped
    for (const id of [
      'formatDialog',
      'conditional',
      'iterative',
      'pasteSpecial',
      'pivotTableDialog',
      'validation',
    ]) {
      expect(features[id], `presets.standard() should NOT activate ${id}`).toBeFalsy();
    }
  });

  it('presets.minimal() keeps only the bare grid surface', async () => {
    sheet = await mountStubSheet({ features: presets.minimal() });
    const { features, host } = sheet.instance;
    // bare chrome must still be present
    expect(host.querySelector('.fc-host__formulabar')).not.toBeNull();
    expect(host.querySelector('.fc-host__grid')).not.toBeNull();
    // all of these should be off
    for (const id of [
      'findReplace',
      'pasteSpecial',
      'quickAnalysis',
      'pivotTableDialog',
      'validation',
      'formatDialog',
      'conditional',
      'iterative',
      'viewToolbar',
    ]) {
      expect(features[id], `presets.minimal() should NOT activate ${id}`).toBeFalsy();
    }
  });
});
