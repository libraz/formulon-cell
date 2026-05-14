import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { presets } from '../../../src/extensions/presets.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

describe('mount/chrome — preset → DOM coverage', () => {
  let sheet: MountedStubSheet;

  afterEach(() => {
    sheet?.dispose();
  });

  describe('presets.full() (default)', () => {
    beforeEach(async () => {
      sheet = await mountStubSheet();
    });

    it('mounts formula bar + status bar + grid', () => {
      expect(sheet.host.querySelector('.fc-host__formulabar')).not.toBeNull();
      expect(sheet.host.querySelector('.fc-host__statusbar')).not.toBeNull();
      expect(sheet.host.querySelector('.fc-host__grid')).not.toBeNull();
    });

    it('mounts sheet tabs + view toolbar', () => {
      expect(sheet.host.querySelector('.fc-host__sheetbar-tabs')).not.toBeNull();
      expect(sheet.host.querySelector('.fc-viewbar')).not.toBeNull();
    });

    it('exposes all default features through `instance.features`', () => {
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
        expect(features[id], `feature ${id} should be active in presets.full()`).toBeTruthy();
      }
    });
  });

  describe('presets.minimal()', () => {
    beforeEach(async () => {
      sheet = await mountStubSheet({ features: presets.minimal() });
    });

    it('still mounts formula bar + status bar + grid for the basic surface', () => {
      expect(sheet.host.querySelector('.fc-host__formulabar')).not.toBeNull();
      expect(sheet.host.querySelector('.fc-host__grid')).not.toBeNull();
    });

    it('does not register the heavyweight chrome features', () => {
      const { features } = sheet.instance;
      for (const id of [
        'findReplace',
        'pasteSpecial',
        'quickAnalysis',
        'pivotTableDialog',
        'validation',
        'formatDialog',
      ]) {
        expect(features[id], `feature ${id} should NOT be active in presets.minimal()`).toBeFalsy();
      }
    });
  });

  describe('presets.standard()', () => {
    beforeEach(async () => {
      sheet = await mountStubSheet({ features: presets.standard() });
    });

    it('keeps clipboard / context menu / find-replace / view toolbar / charts', () => {
      const { features } = sheet.instance;
      expect(features.clipboard).toBeTruthy();
      expect(features.contextMenu).toBeTruthy();
      expect(features.findReplace).toBeTruthy();
      expect(features.viewToolbar).toBeTruthy();
      expect(features.charts).toBeTruthy();
    });

    it('drops the authoring dialogs', () => {
      const { features } = sheet.instance;
      for (const id of [
        'formatDialog',
        'conditional',
        'iterative',
        'pasteSpecial',
        'pivotTableDialog',
        'validation',
        'namedRanges',
        'hyperlink',
      ]) {
        expect(
          features[id],
          `feature ${id} should NOT be active in presets.standard()`,
        ).toBeFalsy();
      }
    });
  });
});

describe('mount/chrome — lifecycle', () => {
  it('dispose() clears the fc-host class and host children', async () => {
    const sheet = await mountStubSheet();
    expect(sheet.host.classList.contains('fc-host')).toBe(true);
    expect(sheet.host.children.length).toBeGreaterThan(0);

    sheet.instance.dispose();

    expect(sheet.host.classList.contains('fc-host')).toBe(false);
    expect(sheet.host.children.length).toBe(0);
    sheet.dispose(); // safe to call again
  });

  it('dispose() is idempotent', async () => {
    const sheet = await mountStubSheet();
    sheet.instance.dispose();
    expect(() => sheet.instance.dispose()).not.toThrow();
    sheet.dispose();
  });
});
