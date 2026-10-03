import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { root, setActive } from './fixtures.js';

describe('attachFormatDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setActive(store, 0, 0);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('keeps Format Cells tabs visually aligned with Japanese Excel 365 desktop', () => {
    const frameCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/frame.css'),
      'utf8',
    );
    const paperThemeCss = readFileSync(join(root, 'src/styles/theme-paper.css'), 'utf8');

    expect(frameCss).toMatch(/\.fc-fmtdlg__panel\s*\{[\s\S]*?border-radius: 4px;/);
    expect(frameCss).toMatch(/\.fc-fmtdlg__preview-cell\s*\{[\s\S]*?border-radius: 3px;/);
    expect(frameCss).toMatch(/\.fc-fmtdlg__tabs\s*\{[\s\S]*?align-self: center;/);
    expect(frameCss).toMatch(/\.fc-fmtdlg__tabs\s*\{[\s\S]*?border-radius: 5px;/);
    expect(frameCss).toMatch(/\.fc-fmtdlg__tab\s*\{[\s\S]*?min-height: 26px;/);
    expect(frameCss).toMatch(
      /\.fc-fmtdlg__close::before,[\s\S]*?\.fc-fmtdlg__close::after\s*\{[\s\S]*?background: currentColor;[\s\S]*?content: "";/,
    );
    expect(frameCss).toMatch(/\.fc-fmtdlg__close::after\s*\{[\s\S]*?rotate\(-45deg\);/);
    expect(frameCss).toMatch(
      /\.fc-fmtdlg__body\s*\{[\s\S]*?border: 1px solid var\(--fc-fmtdlg-rule\);/,
    );
    expect(paperThemeCss).toContain('--fc-fmtdlg-tab-bg: #f3f3f3;');
    expect(paperThemeCss).toContain('--fc-fmtdlg-tab-active-bg: #107c41;');
  });

  it('keeps Format Cells footer buttons and category list close to Japanese Excel 365 desktop', () => {
    const footerCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/footer.css'),
      'utf8',
    );
    const paperThemeCss = readFileSync(join(root, 'src/styles/theme-paper.css'), 'utf8');

    expect(footerCss).toMatch(/\.fc-fmtdlg__footer\s*\{[\s\S]*?gap: 12px;/);
    expect(footerCss).toMatch(/\.fc-fmtdlg__btn\s*\{[\s\S]*?min-height: 24px;/);
    expect(footerCss).toMatch(/\.fc-fmtdlg__btn\s*\{[\s\S]*?border-radius: 7px;/);
    expect(paperThemeCss).toContain('--fc-fmtdlg-btn-bg: #f3f3f3;');
    expect(paperThemeCss).toContain('--fc-fmtdlg-btn-border: transparent;');
    expect(paperThemeCss).toContain('--fc-fmtdlg-btn-primary-bg: #107c41;');
    expect(paperThemeCss).toContain('--fc-fmtdlg-list-border: #86aa8f;');
    expect(paperThemeCss).toContain('--fc-fmtdlg-cat-selected-bg: #107c41;');
  });

  it('keeps Format Cells Alignment tab close to Japanese Excel 365 desktop', () => {
    const alignCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/align-tab.css'),
      'utf8',
    );

    expect(alignCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="align"\]\s*\{[\s\S]*?grid-template-columns: minmax\(292px, 1fr\) 222px;/,
    );
    // The column rule between the alignment rows and the direction dial rides
    // on the dial block itself, so it always ends where the rows do.
    expect(alignCss).toMatch(
      /\.fc-fmtdlg__align-preview\s*\{[\s\S]*?border-left: 1px solid var\(--fc-fmtdlg-rule\);/,
    );
    expect(alignCss).toMatch(
      /\.fc-fmtdlg__align-select-row,\s*\.fc-fmtdlg__text-direction-row\s*\{[\s\S]*?grid-template-columns: 216px;[\s\S]*?min-height: 48px;/,
    );
    // Each left-column row claims its own grid row: auto-placement would push
    // the horizontal-alignment select below the indent field beside it.
    expect(alignCss).toMatch(/\.fc-fmtdlg__align-select-row--h\s*\{\s*grid-row: 1;/);
    expect(alignCss).toMatch(/\.fc-fmtdlg__align-select-row--v\s*\{\s*grid-row: 2;/);
    expect(alignCss).toMatch(/\.fc-fmtdlg__text-direction-row\s*\{\s*grid-row: 3;/);
    expect(alignCss).toMatch(
      /\.fc-fmtdlg__indent-row\s*\{[\s\S]*?grid-row: 1;[\s\S]*?justify-self: end;/,
    );
    expect(alignCss).toMatch(
      /\.fc-fmtdlg__align-preview-box\s*\{[\s\S]*?grid-template-columns: 32px 100px;[\s\S]*?min-height: 152px;/,
    );
    expect(alignCss).toMatch(/\.fc-fmtdlg__align-preview-vertical\s*\{[\s\S]*?height: 150px;/);
  });

  it('keeps Format Cells Border tab close to Japanese Excel 365 desktop', () => {
    const borderCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/borders.css'),
      'utf8',
    );

    expect(borderCss).toMatch(
      /\.fc-fmtdlg__border-visual\s*\{[\s\S]*?padding-top: 12px;[\s\S]*?border-top: 1px solid var\(--fc-fmtdlg-rule\);/,
    );
    expect(borderCss).toMatch(/\.fc-fmtdlg__border-stage\s*\{[\s\S]*?height: 160px;/);
    expect(borderCss).toMatch(
      /\.fc-fmtdlg__border-preview\s*\{[\s\S]*?inset: 28px 52px 28px 54px;/,
    );
    expect(borderCss).toMatch(
      /\.fc-fmtdlg__border-presets\s*\{[\s\S]*?grid-template-columns: repeat\(3, 54px\);[\s\S]*?min-height: 64px;/,
    );
    expect(borderCss).toMatch(
      /\.fc-fmtdlg__btn\.fc-fmtdlg__border-preset\s*\{[\s\S]*?grid-template-rows: 38px 16px;[\s\S]*?min-width: 54px;/,
    );
    expect(borderCss).toMatch(
      /\.fc-fmtdlg__btn\.fc-fmtdlg__border-preset::before\s*\{[\s\S]*?width: 38px;[\s\S]*?height: 38px;/,
    );
    expect(borderCss).toMatch(
      /\.fc-fmtdlg__border-preset--outline::after\s*\{[\s\S]*?border: 2px solid #7f7f7f;/,
    );
    expect(borderCss).toMatch(
      /\.fc-fmtdlg__border-preset--inside::after\s*\{[\s\S]*?linear-gradient\(#7f7f7f 0 0\) 50% 0 \/ 1px 100% no-repeat,/,
    );
    expect(borderCss).toMatch(/\.fc-fmtdlg__border-hit\s*\{[\s\S]*?border-radius: 5px;/);
    expect(borderCss).toMatch(/\.fc-fmtdlg__border-hit--top\s*\{[\s\S]*?top: 4px;/);
    expect(borderCss).toMatch(/\.fc-fmtdlg__border-hit--left\s*\{[\s\S]*?top: 58px;/);

    // Every border-tab item is placed by hand, and both columns are sized so
    // the reset button stays inside the body instead of being clipped.
    const tabsCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/tabs-content.css'),
      'utf8',
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="border"\] \.fc-fmtdlg__border-style-row\s*\{[\s\S]*?grid-column: 2;[\s\S]*?grid-row: 1;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="border"\] \.fc-fmtdlg__border-color-row\s*\{[\s\S]*?grid-column: 2;[\s\S]*?grid-row: 2;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="border"\] \.fc-fmtdlg__border-presets\s*\{[\s\S]*?grid-column: 1;[\s\S]*?grid-row: 5;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="border"\] \.fc-fmtdlg__row > span:first-child\s*\{[\s\S]*?min-width: 72px;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__border-color-row\s*\{[\s\S]*?grid-template-columns: 72px max-content;/,
    );
  });

  it('keeps Format Cells Fill tab close to Japanese Excel 365 desktop', () => {
    const swatchCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/swatches-and-lines.css'),
      'utf8',
    );

    expect(swatchCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="fill"\]\s*\{[\s\S]*?grid-template-columns: minmax\(240px, 1fr\) minmax\(240px, 1fr\);/,
    );
    expect(swatchCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="fill"\]::before\s*\{[\s\S]*?left: 50%;[\s\S]*?background: var\(--fc-fmtdlg-rule\);/,
    );
    expect(swatchCss).toMatch(
      /\.fc-fmtdlg__fill-bg-row\s*\{[\s\S]*?grid-column: 1;[\s\S]*?padding: 28px 0 0 14px;/,
    );
    expect(swatchCss).toMatch(
      /\.fc-fmtdlg__fill-pattern-color-row\s*\{[\s\S]*?grid-column: 2;[\s\S]*?padding: 28px 26px 0 10px;/,
    );
    expect(swatchCss).toMatch(
      /\.fc-fmtdlg__fill-sample\s*\{[\s\S]*?grid-column: 1 \/ -1;[\s\S]*?border-top: 1px solid var\(--fc-fmtdlg-rule\);/,
    );
    expect(swatchCss).toMatch(
      /\.fc-fmtdlg__fill-sample-box\s*\{[\s\S]*?max-width: 510px;[\s\S]*?height: 76px;/,
    );
    // The background-color palette is the shared widget, framed in place.
    expect(swatchCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="fill"\] \.fc-colorpalette\s*\{[\s\S]*?grid-column: 1;[\s\S]*?grid-row: 2;[\s\S]*?border: 1px solid var\(--fc-fmtdlg-input-hover-border\);/,
    );
  });

  it('keeps Format Cells Font tab close to Japanese Excel 365 desktop', () => {
    const tabsCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/tabs-content.css'),
      'utf8',
    );

    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="font"\]\s*\{[\s\S]*?grid-template-columns: 256px 128px 104px;[\s\S]*?gap: 8px 16px;[\s\S]*?padding: 20px 14px 0;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__font-list\s*\{[\s\S]*?height: 110px;[\s\S]*?min-height: 110px;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__font-list--style\s*\{[\s\S]*?grid-column: 2;[\s\S]*?grid-row: 2;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__normal-font\s*\{[\s\S]*?grid-column: 3;[\s\S]*?grid-row: 4;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="font"\]::before\s*\{[\s\S]*?grid-row: 3;[\s\S]*?background: var\(--fc-fmtdlg-rule\);/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__font-preview-box\s*\{[\s\S]*?max-width: 216px;[\s\S]*?border: 0;/,
    );
  });

  it('keeps Format Cells Protection tab close to Japanese Excel 365 desktop', () => {
    const tabsCss = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/tabs-content.css'),
      'utf8',
    );
    const jaStrings = readFileSync(join(root, 'src/i18n/strings/ja.ts'), 'utf8');

    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="protection"\]\s*\{[\s\S]*?padding: 26px 18px 0;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__panel-tab\[data-fc-tab="protection"\] \.fc-fmtdlg__section-title\s*\{[\s\S]*?display: none;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__protection-hint\s*\{[\s\S]*?max-width: 520px;[\s\S]*?line-height: 1\.45;/,
    );
    expect(tabsCss).toMatch(
      /\.fc-fmtdlg__protection-checks\s*\{[\s\S]*?display: grid;[\s\S]*?grid-template-columns: 1fr;/,
    );
    expect(jaStrings).toContain(
      'セルをロックするか、数式を非表示にするには、ワークシートを保護します。',
    );
  });
});
