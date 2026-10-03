import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { menuIconButton, visualMenuTile } from '../../../../src/toolbar/ribbon/menus/general.js';
import { createTextOrientationMenu } from '../../../../src/toolbar/ribbon/menus/text-orientation.js';
import { readCssWithImports } from '../../../helpers/css.js';
import { collectStringLiteralArgs, collectVisualMenuTileIcons, root } from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
  it('expands ribbon menu CSS imports in their declared order', () => {
    const menuCssFiles = [
      'menu-shells.css',
      'menu-command-icons.css',
      'symbol-color-grids.css',
      'paste-review-protection-name-icons.css',
      'visual-galleries.css',
      'conditional-presets.css',
      'style-galleries.css',
      'conditional-menu.css',
      'ribbon-break.css',
      'menu-icon-overlays.css',
    ];
    const menuStylesDir = join(root, 'src/styles/toolbar/ribbon');
    const barrelPath = join(menuStylesDir, 'menus.css');
    const expected = menuCssFiles
      .map((file) => readCssWithImports(join(menuStylesDir, file)))
      .join('');

    expect(readCssWithImports(barrelPath)).toBe(expected);
    expect(readFileSync(barrelPath, 'utf8')).toBe(
      menuCssFiles.map((file) => `@import "./${file}";\n`).join(''),
    );
  });

  it('expands menu command icon CSS parts in their declared order', () => {
    const iconPartFiles = [
      'icon-box.css',
      'page-layout-and-clear.css',
      'fill.css',
      'freeze-panes.css',
      'cells.css',
      'sort-filter.css',
      'find-select.css',
      'formulas.css',
      'data-tools.css',
      'links-pivot.css',
      'text-script.css',
      'add-ins-pdf.css',
      'title-bar-and-symbol.css',
    ];
    const menuStylesDir = join(root, 'src/styles/toolbar/ribbon');
    const barrelPath = join(menuStylesDir, 'menu-command-icons.css');
    const partsDir = join(menuStylesDir, 'menu-command-icons');
    const expected = iconPartFiles
      .map((file) => readFileSync(join(partsDir, file), 'utf8'))
      .join('');

    expect(readCssWithImports(barrelPath)).toBe(expected);
    expect(readFileSync(barrelPath, 'utf8')).toBe(
      iconPartFiles.map((file) => `@import "./menu-command-icons/${file}";\n`).join(''),
    );
  });

  it('embeds Excel-like SVGs for audited cell formatting menu icons', () => {
    for (const iconSlug of [
      'format-dialog',
      'cell-style-new',
      'cell-style-merge',
      'paste-all',
      'paste-formulas',
      'paste-values',
      'paste-formats',
      'paste-transpose',
      'paste-special',
      'fill-down',
      'fill-right',
      'clear-all',
      'clear-formats',
      'sort-asc',
      'sort-desc',
      'filter-toggle',
      'find',
      'find-formulas',
      'merge',
      'freeze-panes',
      'freeze-row',
      'freeze-col',
      'insert-sheet',
      'delete-sheet',
      'format-row-height',
      'format-col-width',
      'format-lock',
      'format-protect',
      'go-to',
      'go-to-special',
      'remove-duplicates',
      'name-manager',
      'text-column-comma',
      'print-area-set',
      'break-page',
      'bring-forward',
      'send-backward',
      'pivot-range',
      'pivot-recommended',
      'pivot-existing-sheet',
      'defined-name-manager',
      'defined-name-create-top',
      'defined-name-create-bottom',
      'defined-name-create-left',
      'defined-name-create-right',
      'link-edit',
      'validation-settings',
      'validation-circle',
      'validation-clear-circles',
      'validation-clear-rules',
      'script-custom',
      'addin-get',
      'pdf-create',
      'watch-open',
      'comment-delete',
      'protect-sheet',
      'autosum-sum',
      'error-checking',
      'trace-error',
      'currency-yen',
      'currency-dollar',
      'new-table-style',
      'pivot-style-new',
      'title-save',
      'title-save-as',
      'title-autosave',
      'title-comments',
      'title-share',
    ]) {
      const button = menuIconButton('セルの書式設定...', 'cellFormat', 'dialog', iconSlug);
      const icon = button.querySelector('.fc-tb__menu-icon');

      expect(icon?.classList.contains('fc-tb__menu-icon--svg')).toBe(true);
      expect(icon?.querySelector('.fc-tb__menu-icon-svg')).toBeTruthy();
      expect(icon?.querySelectorAll('path').length).toBeGreaterThan(0);
    }
  });

  it('keeps every real menuIconButton icon slug connected to an Excel-like SVG', () => {
    const slugs = collectStringLiteralArgs('menuIconButton', 3);
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(slugs.length).toBeGreaterThan(100);
    for (const iconSlug of slugs) {
      const button = menuIconButton(iconSlug, 'auditAction', iconSlug, iconSlug);
      const icon = button.querySelector('.fc-tb__menu-icon');

      expect(icon?.classList.contains('fc-tb__menu-icon--svg'), iconSlug).toBe(true);
      expect(icon?.querySelector('.fc-tb__menu-icon-svg'), iconSlug).toBeTruthy();
      expect(icon?.querySelectorAll('path').length, iconSlug).toBeGreaterThan(0);
    }
    expect(menusCss).toMatch(
      /\.fc-tb__menu-item__icon-spacer\s*\{[\s\S]*?flex: 0 0 18px;[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon\s*\{[\s\S]*?flex: 0 0 18px;[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps edit/save fallback menu glyphs as pencil overlays, not placeholder text', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--format-rename-sheet::after',
      '.fc-tb__menu-icon--link-edit::after',
      '.fc-tb__menu-icon--title-save-as::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}\\s*\\{[\\s\\S]*?width: 8px;[\\s\\S]*?height: 3px;[\\s\\S]*?background: #185abd;[\\s\\S]*?box-shadow: -2px 0 0 #f4b183;[\\s\\S]*?content: "";[\\s\\S]*?transform: rotate\\(-35deg\\);`,
        ),
      );
    }
    expect(menusCss).not.toContain('content: "I"');
  });

  it('renders delete and clear fallback menu glyphs as vector crosses, not lowercase text', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--clear::after',
      '.fc-tb__menu-icon--delete-sheet::after',
      '.fc-tb__menu-icon--filter-clear::after',
      '.fc-tb__menu-icon--validation-clear-rules::after',
      '.fc-tb__menu-icon--link-clear::after',
      '.fc-tb__menu-icon--comment-delete::after',
      '.fc-tb__menu-icon--protect-clear-ranges::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?\\{[\\s\\S]*?width: 10px;[\\s\\S]*?height: 10px;[\\s\\S]*?linear-gradient\\(45deg,[\\s\\S]*?#a4262c[\\s\\S]*?linear-gradient\\(135deg,[\\s\\S]*?#a4262c[\\s\\S]*?content: "";`,
        ),
      );
    }
    expect(menusCss).not.toContain('content: "x"');
  });

  it('renders custom sort fallback glyph as fixed arrows, not a font symbol', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--sort-asc::after',
      '.fc-tb__menu-icon--sort-desc::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?width: 13px;[\\s\\S]*?height: 13px;[\\s\\S]*?background-image: url\\("data:image/svg\\+xml,[\\s\\S]*?fill='%23185abd'[\\s\\S]*?stroke='%23107c41'[\\s\\S]*?content: "";`,
        ),
      );
    }
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--sort-custom::after\s*\{[\s\S]*?width: 13px;[\s\S]*?height: 13px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%23185abd'[\s\S]*?background-size: 13px 13px;[\s\S]*?content: "";/,
    );
    const sortCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--sort-asc::after'),
      menusCss.indexOf('.fc-tb__menu-icon--sort-custom::after'),
    );
    expect(sortCss).not.toContain('content: "A"');
    expect(sortCss).not.toContain('content: "Z"');
    expect(menusCss).not.toContain('content: "⇅"');
  });

  it('renders fill direction fallback glyphs as fixed arrows, not font symbols', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--fill-down::after',
      '.fc-tb__menu-icon--fill-up::after',
      '.fc-tb__menu-icon--fill-right::after',
      '.fc-tb__menu-icon--fill-left::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}\\s*\\{[\\s\\S]*?background-image: url\\("data:image/svg\\+xml,[\\s\\S]*?stroke='%23107c41'[\\s\\S]*?content: "";`,
        ),
      );
    }
  });

  it('renders sheet move fallback glyphs as fixed arrows, not font symbols', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--format-move-left::after,[\s\S]*?\.fc-tb__menu-icon--format-move-right::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?background-size: 12px 12px;[\s\S]*?content: "";/,
    );
    for (const selector of [
      '.fc-tb__menu-icon--format-move-left::after',
      '.fc-tb__menu-icon--format-move-right::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}\\s*\\{[\\s\\S]*?background-image: url\\("data:image/svg\\+xml,[\\s\\S]*?stroke='%23107c41'`,
        ),
      );
    }
  });

  it('renders Go To fallback glyph as a fixed arrow, not a font symbol', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--go-to::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%23185abd'[\s\S]*?background-size: 12px 12px;[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "➜"');
  });

  it('renders filter value and advanced glyphs as fixed marks, not font symbols', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--filter-by-value::after\s*\{[\s\S]*?width: 10px;[\s\S]*?height: 8px;[\s\S]*?linear-gradient\(#107c41 0 0\) 1px 2px \/ 8px 2px no-repeat,[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--filter-advanced::after\s*\{[\s\S]*?width: 11px;[\s\S]*?height: 5px;[\s\S]*?radial-gradient\(circle at 2px 50%, #605e5c[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "="');
    expect(menusCss).not.toContain('content: "⋯"');
  });

  it('renders Text to Columns delimiter glyphs as fixed marks, not font text', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--text-column-comma::after',
      '.fc-tb__menu-icon--text-column-tab::after',
      '.fc-tb__menu-icon--text-column-semicolon::after',
      '.fc-tb__menu-icon--text-column-space::after',
      '.fc-tb__menu-icon--text-column-custom::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(`${escaped}\\s*\\{[\\s\\S]*?width: 12px;[\\s\\S]*?content: "";`),
      );
    }

    const textColumnCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--text-column-comma::before'),
      menusCss.indexOf('.fc-tb__menu-icon--link-clear::before'),
    );
    for (const glyph of ['","', '"Tab"', '";"', '"␠"', '"…"']) {
      expect(textColumnCss).not.toContain(`content: ${glyph}`);
    }
  });

  it('renders Remove Duplicates fallback glyph as overlapped records, not text', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--remove-duplicates::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?linear-gradient\(#ffffff 0 0\) 3px 1px \/ 7px 7px no-repeat,[\s\S]*?border: 1px solid #a4262c;[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "2"');
  });

  it('renders formula and calculation fallback glyphs as fixed marks, not font symbols', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--find-formulas::after\s*\{[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%238764b8'[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--calc-auto::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%23107c41'[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--calc-auto-no-table::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?linear-gradient\(45deg,[\s\S]*?#a4262c[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--calc-manual::after\s*\{[\s\S]*?width: 11px;[\s\S]*?height: 11px;[\s\S]*?linear-gradient\(#605e5c 0 0\) 2px 4\.5px \/ 7px 2px no-repeat,[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--calc-sheet::after\s*\{[\s\S]*?width: 11px;[\s\S]*?height: 11px;[\s\S]*?linear-gradient\(#185abd 0 0\) 0 50% \/ 100% 1\.4px no-repeat,[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--calc-iterative::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 10px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%238764b8'[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "ƒ"');
    expect(menusCss).not.toContain('content: "▦"');
    expect(menusCss).not.toContain('content: "∞"');
    const calcCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--calc-auto-no-table::before'),
      menusCss.indexOf('.fc-tb__menu-item[role="menuitemradio"][aria-checked="true"]'),
    );
    expect(calcCss).not.toContain('content: "A"');
    expect(calcCss).not.toContain('content: "A*"');
    expect(calcCss).not.toContain('content: "M"');
  });

  it('renders arrange front/back badges as vector plates, not numeric text', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--bring-front::after,[\s\S]*?\.fc-tb__menu-icon--send-back::after\s*\{[\s\S]*?width: 7px;[\s\S]*?height: 7px;[\s\S]*?border: 1px solid #0b5a2f;[\s\S]*?background: #107c41;[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "1"');
  });

  it('renders conditional-formatting symbol icons as vector marks, not text glyphs', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--check-green::before\s*\{[\s\S]*?border-bottom: 2px solid currentColor;[\s\S]*?border-left: 2px solid currentColor;[\s\S]*?transform: rotate\(-45deg\);/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--bang-yellow::before\s*\{[\s\S]*?width: 2px;[\s\S]*?height: 7px;[\s\S]*?background: currentColor;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--x-red::before,[\s\S]*?\.fc-tb__cf-icon--symbol\.fc-tb__cf-icon--x-red::after\s*\{[\s\S]*?width: 2px;[\s\S]*?height: 10px;[\s\S]*?background: currentColor;/,
    );
    const symbolCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__cf-icon--symbol::before'),
      menusCss.indexOf('.fc-tb__cf-icon--flag::before'),
    );
    expect(symbolCss).not.toContain('content: "✓"');
    expect(symbolCss).not.toContain('content: "!"');
    expect(symbolCss).not.toContain('content: "×"');
  });

  it('renders script custom and symbol more glyphs as SVG marks, not font text', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--script-uppercase::after',
      '.fc-tb__menu-icon--script-lowercase::after',
      '.fc-tb__menu-icon--script-trim::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?background-image: url\\("data:image/svg\\+xml,[\\s\\S]*?fill='%23`,
        ),
      );
    }
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--script-custom::after\s*\{[\s\S]*?width: 13px;[\s\S]*?height: 12px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%238764b8'[\s\S]*?background-size: 13px 12px;[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--symbol-more::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%23185abd'[\s\S]*?background-size: 12px 12px;[\s\S]*?content: "";/,
    );
    const scriptCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--script-clear::after'),
      menusCss.indexOf('.fc-tb__menu-icon--addin-get::before'),
    );
    expect(scriptCss).not.toContain('content: "A"');
    expect(scriptCss).not.toContain('content: "a"');
    expect(scriptCss).not.toContain('content: "T"');
    expect(menusCss).not.toContain('content: "{}"');
    expect(menusCss).not.toContain('content: "Ω"');
  });

  it('renders Watch Window open glyph as an eye mark, not a W character', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--watch-open::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 9px;[\s\S]*?radial-gradient\(circle at 50% 50%, #185abd[\s\S]*?radial-gradient\(ellipse at 50% 50%[\s\S]*?content: "";/,
    );
    const watchOpenCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--watch-open::after'),
      menusCss.indexOf('.fc-tb__menu-icon--protect-allow-ranges::before'),
    );
    expect(watchOpenCss).not.toContain('content: "W"');
  });

  it('renders name and formula-use badges as fixed marks, not N or fx text', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--name-manager::after',
      '.fc-tb__menu-icon--defined-name-define::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?width: 12px;[\\s\\S]*?height: 10px;[\\s\\S]*?background-image: url\\("data:image/svg\\+xml,[\\s\\S]*?stroke='%23107c41'[\\s\\S]*?content: "";`,
        ),
      );
    }

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--defined-name-manager::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 11px;[\s\S]*?border: 1px solid #185abd;[\s\S]*?linear-gradient\(#185abd 0 0\)[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--defined-name-use::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 11px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%238764b8'[\s\S]*?content: "";/,
    );
    for (const selector of [
      '.fc-tb__menu-icon--defined-name-create-top::after',
      '.fc-tb__menu-icon--defined-name-create-bottom::after',
      '.fc-tb__menu-icon--defined-name-create-left::after',
      '.fc-tb__menu-icon--defined-name-create-right::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?background-image: url\\("data:image/svg\\+xml,[\\s\\S]*?stroke='%23107c41'[\\s\\S]*?`,
        ),
      );
    }

    const nameCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--name-manager::after'),
      menusCss.indexOf('.fc-tb__menu-icon--find-comments::before'),
    );
    const definedNameCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--defined-name-define::after'),
      menusCss.indexOf('.fc-tb__menu-sep'),
    );
    expect(`${nameCss}\n${definedNameCss}`).not.toContain('content: "N"');
    expect(definedNameCss).not.toContain('content: "fx"');
    expect(definedNameCss).not.toContain('content: "T"');
    expect(definedNameCss).not.toContain('content: "B"');
    expect(definedNameCss).not.toContain('content: "L"');
    expect(definedNameCss).not.toContain('content: "R"');
  });

  it('renders Format Cells dialog badge as an edit mark, not an A character', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--format-dialog::after\s*\{[\s\S]*?width: 11px;[\s\S]*?height: 8px;[\s\S]*?background: #185abd;[\s\S]*?box-shadow: -2px 0 0 #f4b183;[\s\S]*?content: "";[\s\S]*?transform: rotate\(-35deg\);/,
    );
    const formatDialogCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--format-dialog::after'),
      menusCss.indexOf('.fc-tb__menu-icon--format-row-height::after'),
    );
    expect(formatDialogCss).not.toContain('content: "A"');
  });

  it('renders My Add-ins badge as add-in tiles, not an M character', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--addin-my::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?linear-gradient\(#185abd 0 0\)[\s\S]*?linear-gradient\(#8764b8 0 0\)[\s\S]*?content: "";/,
    );
    const addinMyCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--addin-my::after'),
      menusCss.indexOf('.fc-tb__menu-icon--addin-manage::after'),
    );
    expect(addinMyCss).not.toContain('content: "M"');
  });

  it('renders PivotTable existing sheet badge as a target cell, not a D character', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--pivot-existing-sheet::after\s*\{[\s\S]*?width: 11px;[\s\S]*?height: 11px;[\s\S]*?border: 2px solid #185abd;[\s\S]*?linear-gradient\(#185abd 0 0\) 3px 3px \/ 3px 3px no-repeat,[\s\S]*?content: "";/,
    );
    const pivotExistingCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--pivot-existing-sheet::after'),
      menusCss.indexOf('.fc-tb__menu-icon--script-clear::before'),
    );
    expect(pivotExistingCss).not.toContain('content: "D"');
  });

  it('renders add, launch, and settings fallback menu glyphs as vector overlays', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--insert-sheet::after',
      '.fc-tb__menu-icon--pivot-new-sheet::after',
      '.fc-tb__menu-icon--fill-series::after',
      '.fc-tb__menu-icon--addin-get::after',
      '.fc-tb__cellstyle-footer:not(.fc-tb__menu-item--iconic)::before',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?width: 10px;[\\s\\S]*?height: 10px;[\\s\\S]*?linear-gradient\\(#107c41 0 0\\)[\\s\\S]*?linear-gradient\\(90deg, #107c41 0 0\\)[\\s\\S]*?content: "";`,
        ),
      );
    }

    for (const selector of [
      '.fc-tb__menu-icon--link-open::after',
      '.fc-tb__menu-icon--pdf-share::after',
      '.fc-tb__menu-icon--title-share::after',
      '.fc-tb__menu-icon--trace-error::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?width: 10px;[\\s\\S]*?height: 10px;[\\s\\S]*?linear-gradient\\(#185abd 0 0\\)[\\s\\S]*?linear-gradient\\(45deg,[\\s\\S]*?#185abd[\\s\\S]*?content: "";`,
        ),
      );
    }

    for (const selector of [
      '.fc-tb__menu-icon--addin-manage::after',
      '.fc-tb__menu-icon--pdf-preferences::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?width: 12px;[\\s\\S]*?height: 12px;[\\s\\S]*?radial-gradient\\(circle,[\\s\\S]*?conic-gradient\\([\\s\\S]*?content: "";`,
        ),
      );
    }
    expect(menusCss).not.toContain('content: "+"');
    expect(menusCss).not.toContain('content: "↗"');
    expect(menusCss).not.toContain('content: "⚙"');
  });

  it('renders star fallback menu glyphs as filled star shapes, not font characters', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--go-to-special::after,[\s\S]*?\.fc-tb__menu-icon--pivot-recommended::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?background: #d83b01;[\s\S]*?clip-path: polygon\([\s\S]*?50% 0,[\s\S]*?content: "";/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--go-to-special::after\s*\{[\s\S]*?background: #8764b8;/,
    );
    expect(menusCss).not.toContain('content: "★"');
  });

  it('uses semantic SVGs for underline variant menu icons', () => {
    for (const iconSlug of ['underline-single', 'underline-double']) {
      const button = menuIconButton('下線', 'underlineAction', 'single', iconSlug);
      const icon = button.querySelector('.fc-tb__menu-icon');

      expect(icon?.classList.contains('fc-tb__menu-icon--svg')).toBe(true);
      expect(icon?.querySelector('.fc-tb__menu-icon-svg')).toBeTruthy();
      // The rule under the letterform is the underline itself, so it carries
      // the ink colour rather than an accent.
      expect(icon?.querySelector('path[fill="#1f1f1f"]')).toBeTruthy();
    }
  });

  it('renders checked, reapply, and warning fallback menu glyphs as vector overlays', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));

    for (const selector of [
      '.fc-tb__menu-icon--format-unhide-sheet::after',
      '.fc-tb__menu-icon--find-validation::after',
      '.fc-tb__menu-icon--ignore-error::after',
      '.fc-tb__menu-icon--validation-settings::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?width: 9px;[\\s\\S]*?height: 5px;[\\s\\S]*?border-bottom: 2px solid #107c41;[\\s\\S]*?border-left: 2px solid #107c41;[\\s\\S]*?content: "";[\\s\\S]*?transform: rotate\\(-45deg\\);`,
        ),
      );
    }
    expect(menusCss).not.toContain('content: "✓"');

    for (const selector of [
      '.fc-tb__menu-icon--filter-reapply::after',
      '.fc-tb__menu-icon--calc-now::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(
          `${escaped}[\\s\\S]*?width: 11px;[\\s\\S]*?height: 11px;[\\s\\S]*?border: 2px solid #107c41;[\\s\\S]*?border-left-color: transparent;[\\s\\S]*?content: "";`,
        ),
      );
    }
    expect(menusCss).not.toContain('content: "↻"');

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--error-checking::after[\s\S]*?\{[\s\S]*?width: 9px;[\s\S]*?height: 12px;[\s\S]*?radial-gradient\(circle at 5px 11px,[\s\S]*?content: "";/,
    );
  });

  it('uses semantic SVGs for visual chart, picture, shape, and screenshot tiles', () => {
    for (const iconSlug of [
      'chart-column',
      'chart-bar',
      'chart-line',
      'chart-area',
      'chart-pie',
      'chart-scatter',
      'chart-recommended',
      'device-picture',
      'online-picture',
      'stock-picture',
      'shape-line',
      'shape-arrow',
      'shape-rectangle',
      'shape-rounded-rectangle',
      'shape-oval',
      'shape-triangle',
      'shape-diamond',
      'screenshot-window',
      'screen-clipping',
      'theme-light',
      'theme-dark',
      'theme-contrast',
    ]) {
      const button = visualMenuTile({
        label: iconSlug,
        attr: 'visualAction',
        value: iconSlug,
        icon: iconSlug,
      });
      const icon = button.querySelector('.fc-tb__visual-tile__icon');

      expect(icon?.classList.contains('fc-tb__visual-tile__icon--svg')).toBe(true);
      expect(icon?.querySelector('.fc-tb__visual-tile__icon-svg')).toBeTruthy();
      expect(icon?.querySelectorAll('path').length).toBeGreaterThan(0);
    }
  });

  it('keeps every real visualMenuTile icon slug connected to a semantic SVG', () => {
    const slugs = collectVisualMenuTileIcons();

    expect(slugs.length).toBeGreaterThan(20);
    for (const iconSlug of slugs) {
      const button = visualMenuTile({
        label: iconSlug,
        attr: 'visualAction',
        value: iconSlug,
        icon: iconSlug,
      });
      const icon = button.querySelector('.fc-tb__visual-tile__icon');

      expect(icon?.classList.contains('fc-tb__visual-tile__icon--svg'), iconSlug).toBe(true);
      expect(icon?.querySelector('.fc-tb__visual-tile__icon-svg'), iconSlug).toBeTruthy();
      expect(icon?.querySelectorAll('path').length, iconSlug).toBeGreaterThan(0);
    }
  });

  it('renders Text Orientation menu previews as colored semantic paths', () => {
    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
    const menu = createTextOrientationMenu({
      orientationAngleCounterclockwise: '左回りに回転',
      orientationAngleClockwise: '右回りに回転',
      orientationVerticalText: '縦書き',
      orientationRotateTextUp: '上へ回転',
      orientationRotateTextDown: '下へ回転',
      orientationFormatAlignment: 'セルの配置の設定',
    } as Parameters<typeof createTextOrientationMenu>[0]);
    const previews = Array.from(
      menu.querySelectorAll<SVGSVGElement>('.fc-tb__text-orientation-preview'),
    );
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-text-orientation]'));

    expect(menu.id).toBe('menu-text-orientation');
    expect(items.map((item) => item.dataset.textOrientation)).toEqual([
      'ccw',
      'cw',
      'vertical',
      'up',
      'down',
      'format',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '左回りに回転',
      '右回りに回転',
      '縦書き',
      '上へ回転',
      '下へ回転',
      'セルの配置の設定',
    ]);
    expect(previews).toHaveLength(6);
    for (const preview of previews) {
      expect(preview.querySelector('text')).toBeNull();
      expect(preview.querySelectorAll('path').length).toBeGreaterThan(2);
      expect(preview.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    }
    expect(previews.some((preview) => preview.querySelector('path[stroke="#2f75b5"]'))).toBe(true);
    expect(menusCss).toMatch(/#menu-text-orientation\s*\{[\s\S]*?min-width: 168px;/);
    expect(menusCss).toMatch(
      /#menu-text-orientation \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-text-orientation \.fc-tb__text-orientation-preview\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });
});
