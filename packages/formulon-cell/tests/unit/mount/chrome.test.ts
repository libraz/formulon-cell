import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { presets } from '../../../src/extensions/presets.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

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

describe('mount/chrome — Excel-like formula bar surface', () => {
  it('keeps the formula bar compact with neutral name-box popup states', () => {
    const css = readFileSync(join(root, 'src/styles/core/surface/formulabar.css'), 'utf8');

    expect(css).toContain('--fc-formulabar-namebox-width: 86px;');
    expect(css).toContain('--fc-formulabar-height: 28px;');
    expect(css).toContain('--fc-formulabar-control-height: 20px;');
    expect(css).toContain('--fc-formulabar-action-size: 22px;');
    expect(css).toContain('gap: 3px;');
    expect(css).toContain('--fc-formulabar-fx-width: 26px;');
    expect(css).toMatch(/\.fc-host__formulabar\s*\{[\s\S]*?padding: 3px 5px;/);
    expect(css).toMatch(
      /\.fc-host__formulabar-fx\s*\{[\s\S]*?font-style: italic;[\s\S]*?border-right: var\(--fc-hairline\) solid var\(--fc-rule\);/,
    );
    expect(css).toMatch(
      /\.fc-host__formulabar-input\s*\{[\s\S]*?min-height: var\(--fc-formulabar-control-height\);[\s\S]*?border-radius: 2px;/,
    );
    const nameBoxCss = css.slice(
      css.indexOf('.fc-host__formulabar-tag {'),
      css.indexOf('.fc-host__formulabar-tag:focus'),
    );
    expect(nameBoxCss).toContain('padding: 1px 17px 1px 6px;');
    expect(nameBoxCss).toContain('linear-gradient(45deg, transparent 50%, var(--fc-fg-mute) 50%)');
    expect(nameBoxCss).toContain('background-size: 4px 4px;');
    expect(css).toContain('box-shadow: inset 0 0 0 1px var(--fc-accent);');
    expect(css).toMatch(
      /\.fc-host__formulabar-expand\s*\{[\s\S]*?border: 0;[\s\S]*?border-left: var\(--fc-hairline\) solid var\(--fc-rule\);[\s\S]*?border-radius: 0;/,
    );
    expect(css).toMatch(
      /\.fc-host__formulabar-expand:hover\s*\{[\s\S]*?background: var\(--fc-bg-hover\);[\s\S]*?color: var\(--fc-fg\);/,
    );
    expect(css).toContain('.fc-host__formulabar-input:focus');
    expect(css).toMatch(
      /\.fc-namebox-menu__item:hover,[\s\S]*?\.fc-namebox-menu__item:focus-visible\s*\{[\s\S]*?background: var\(--fc-bg-hover\);/,
    );
    expect(css).not.toContain(
      '.fc-namebox-menu__item:focus-visible {\n    background: var(--fc-accent-soft);',
    );
  });
});

describe('mount/chrome — Excel-like bottom chrome surface', () => {
  it('keeps sheet tabs compact with neutral hover and green active text', () => {
    const css = readFileSync(join(root, 'src/styles/core/surface/sheetbar.css'), 'utf8');

    expect(css).toContain('--fc-sheetbar-height: 28px;');
    expect(css).toContain('--fc-sheetbar-tab-height: 26px;');
    expect(css).toContain('--fc-sheetbar-tab-min-width: 78px;');
    expect(css).toMatch(
      /\.fc-host__sheetbar-tab\s*\{[\s\S]*?display: inline-flex;[\s\S]*?align-items: center;[\s\S]*?justify-content: center;/,
    );
    expect(css).toContain('background: var(--fc-bg-hover);');
    expect(css).toContain('color: var(--fc-accent);');
    expect(css).toContain('font-weight: 400;');
    expect(css).toMatch(
      /\.fc-host__sheetbar-tab\[aria-selected="true"\]::before\s*\{[\s\S]*?height: 3px;/,
    );
    expect(css).toMatch(
      /\.fc-host__sheetbar-add\s*\{[\s\S]*?border: var\(--fc-hairline\) solid var\(--fc-statusbar-control, var\(--fc-rule\)\);[\s\S]*?border-radius: 50%;/,
    );
    expect(css).toMatch(
      /\.fc-host__sheetbar-add \.fc-host__icon\s*\{[\s\S]*?width: 14px;[\s\S]*?height: 14px;/,
    );
    expect(css).toMatch(
      /\.fc-host__sheetbar-rename\s*\{[\s\S]*?height: var\(--fc-sheetbar-rename-height\);[\s\S]*?border: 1px solid var\(--fc-accent\);[\s\S]*?border-radius: 2px;[\s\S]*?box-shadow: none;/,
    );
    expect(css).toMatch(
      /\.fc-host__sheetbar-rename\[aria-invalid="true"\]\s*\{[\s\S]*?border-color: var\(--fc-cell-error-fg\);[\s\S]*?box-shadow: none;/,
    );
    expect(css).not.toContain('background: var(--fc-accent-soft);');
  });

  it('keeps the status bar on the neutral Excel 365 desktop surface', () => {
    const theme = readFileSync(join(root, 'src/styles/theme-paper.css'), 'utf8');
    const css = readFileSync(join(root, 'src/styles/core/surface/statusbar.css'), 'utf8');

    expect(theme).toContain('--fc-statusbar-bg: #f3f2f1;');
    expect(theme).toContain('--fc-statusbar-border: #d9d9d9;');
    expect(theme).toContain('--fc-statusbar-fg: #605e5c;');
    expect(theme).toContain('--fc-statusbar-fg-strong: #201f1e;');
    expect(theme).toContain('--fc-statusbar-control: #c8c6c4;');
    expect(css).toContain('min-height: 24px;');
    expect(css).toContain('padding: 2px 10px;');
    expect(css).toMatch(
      /\.fc-host__statusbar\s*\{[\s\S]*?letter-spacing: 0;[\s\S]*?text-transform: none;/,
    );
    expect(css).toMatch(
      /\.fc-host__statusbar-view\s*\{[\s\S]*?width: 24px;[\s\S]*?height: 22px;[\s\S]*?border-radius: 3px;/,
    );
    expect(css).toMatch(
      /\.fc-host__statusbar-viewicon\s*\{[\s\S]*?width: 16px;[\s\S]*?height: 16px;/,
    );
    expect(css).toContain('0 5px / 100% 1px no-repeat');
    expect(css).toContain('5px 0 / 1px 100% no-repeat');
    expect(css).toContain('0 7px / 100% 1px no-repeat');
    expect(css).toContain('border-color: var(--fc-statusbar-control, var(--fc-rule));');
    expect(css).not.toContain('--fc-statusbar-bg: var(--fc-accent);');
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

  it('keeps host chrome buttons on the shared host button helper', () => {
    const source = readFileSync(join(root, 'src/mount/chrome.ts'), 'utf8');
    expect(source).toContain('createHostButton({');
    expect(source).not.toContain("document.createElement('button')");
  });
});
