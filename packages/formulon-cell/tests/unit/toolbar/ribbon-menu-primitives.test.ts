import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
import { describe, expect, it, vi } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import {
  focusMenuItem,
  handleMenuKeydown,
  prepareMenu,
  projectDisabledReason,
  projectDisabledState,
} from '../../../src/toolbar/menu-a11y.js';
import {
  RIBBON_BORDERS_MENU_ID,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
} from '../../../src/toolbar/ribbon/activation.js';
import {
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS,
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS,
  DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS,
  type DynamicDropdownsCtx,
} from '../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import { createBordersMenu } from '../../../src/toolbar/ribbon/menus/borders.js';
import { createConditionalMenu } from '../../../src/toolbar/ribbon/menus/conditional.js';
import { createFormulasMenuFactories } from '../../../src/toolbar/ribbon/menus/formulas.js';
import {
  colorSwatchButton,
  colorSwatchGrid,
  createMenu,
  createMenuButton,
  createSubmenu,
  menuIconButton,
  menuIconSpacer,
  menuLabeledGrid,
  menuPresetButton,
  menuSectionHeader,
  menuSeparator,
  menuSubmenuTrigger,
  menuTextChip,
  submenuItemText,
  symbolMenuGrid,
  symbolMenuTile,
  visualMenuGrid,
  visualMenuTile,
} from '../../../src/toolbar/ribbon/menus/general.js';
import { createHomeMenuFactories } from '../../../src/toolbar/ribbon/menus/home.js';
import { createStylesMenuFactories } from '../../../src/toolbar/ribbon/menus/styles.js';
import { createTextOrientationMenu } from '../../../src/toolbar/ribbon/menus/text-orientation.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const menusDir = join(root, 'src/toolbar/ribbon/menus');
const ribbonDir = join(root, 'src/toolbar/ribbon');
const mountDir = join(root, 'src/mount');
const disabledStateAuditDirs = ['src/interact', 'src/mount', 'src/toolbar', 'src/components'];

const menuSources = (): { name: string; source: string }[] =>
  readdirSync(menusDir)
    .filter((name) => name.endsWith('.ts'))
    .map((name) => ({
      name,
      source: readFileSync(join(menusDir, name), 'utf8'),
    }));

const menuConsumerSources = (): { name: string; source: string }[] => [
  ...menuSources(),
  {
    name: 'backstage-title.ts',
    source: readFileSync(join(ribbonDir, 'backstage-title.ts'), 'utf8'),
  },
];

const sourceFilesUnder = (path: string): string[] => {
  const absolutePath = join(root, path);
  const files: string[] = [];
  for (const entry of readdirSync(absolutePath, { withFileTypes: true })) {
    const entryPath = `${path}/${entry.name}`;
    if (entry.isDirectory()) {
      files.push(...sourceFilesUnder(entryPath));
    } else if (entry.name.endsWith('.ts')) {
      files.push(entryPath);
    }
  }
  return files.sort();
};

const collectStringLiteralArgs = (callName: string, argIndex: number): string[] => {
  const values = new Set<string>();
  for (const { name, source } of menuConsumerSources()) {
    if (name === 'general.ts') continue;
    const file = ts.createSourceFile(name, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
    const visit = (node: ts.Node): void => {
      if (
        ts.isCallExpression(node) &&
        ts.isIdentifier(node.expression) &&
        node.expression.text === callName
      ) {
        const arg = node.arguments[argIndex];
        if (arg && ts.isStringLiteralLike(arg)) values.add(arg.text);
      }
      ts.forEachChild(node, visit);
    };
    visit(file);
  }
  return [...values].sort();
};

const collectVisualMenuTileIcons = (): string[] => {
  const values = new Set<string>();
  const collectIconFromObject = (opts: ts.ObjectLiteralExpression): void => {
    const iconProp = opts.properties.find(
      (prop): prop is ts.PropertyAssignment =>
        ts.isPropertyAssignment(prop) &&
        ts.isIdentifier(prop.name) &&
        prop.name.text === 'icon' &&
        ts.isStringLiteralLike(prop.initializer),
    );
    if (iconProp && ts.isStringLiteralLike(iconProp.initializer)) {
      values.add(iconProp.initializer.text);
    }
  };

  for (const { name, source } of menuConsumerSources()) {
    if (name === 'general.ts') continue;
    const file = ts.createSourceFile(name, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
    const visit = (node: ts.Node): void => {
      if (!ts.isCallExpression(node) || !ts.isIdentifier(node.expression)) {
        ts.forEachChild(node, visit);
        return;
      }

      if (node.expression.text === 'visualMenuTile') {
        const opts = node.arguments[0];
        if (opts && ts.isObjectLiteralExpression(opts)) {
          collectIconFromObject(opts);
        }
      }
      if (node.expression.text === 'visualMenuTileGrid') {
        const tiles = node.arguments[1];
        if (tiles && ts.isArrayLiteralExpression(tiles)) {
          for (const element of tiles.elements) {
            if (ts.isObjectLiteralExpression(element)) collectIconFromObject(element);
          }
        }
      }
      ts.forEachChild(node, visit);
    };
    visit(file);
  }
  return [...values].sort();
};

describe('toolbar/ribbon menu primitives', () => {
  const sourcesOutsidePrimitives = (): { name: string; source: string }[] =>
    menuSources().filter(({ name }) => name !== 'general.ts');

  it('keeps preset menu row DOM centralized in menuPresetButton', () => {
    const directPresetRows = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes('fc-tb__menu-item fc-tb__menu-item--preset'))
      .map(({ name }) => name);

    expect(directPresetRows).toEqual([]);
  });

  it('keeps iconic menu row DOM centralized in menuIconButton', () => {
    const directIconicRows = sourcesOutsidePrimitives()
      .filter(
        ({ source }) =>
          source.includes('fc-tb__menu-item fc-tb__menu-item--iconic') ||
          source.includes('fc-tb__menu-icon fc-tb__menu-icon--'),
      )
      .map(({ name }) => name);

    expect(directIconicRows).toEqual([]);
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
      expect(icon?.querySelectorAll('path').length).toBeGreaterThan(1);
    }
  });

  it('keeps every real menuIconButton icon slug connected to an Excel-like SVG', () => {
    const slugs = collectStringLiteralArgs('menuIconButton', 3);
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--go-to::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?background-image: url\("data:image\/svg\+xml,[\s\S]*?stroke='%23185abd'[\s\S]*?background-size: 12px 12px;[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "➜"');
  });

  it('renders filter value and advanced glyphs as fixed marks, not font symbols', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--remove-duplicates::after\s*\{[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?linear-gradient\(#ffffff 0 0\) 3px 1px \/ 7px 7px no-repeat,[\s\S]*?border: 1px solid #a4262c;[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "2"');
  });

  it('renders formula and calculation fallback glyphs as fixed marks, not font symbols', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

    expect(menusCss).toMatch(
      /\.fc-tb__menu-icon--bring-front::after,[\s\S]*?\.fc-tb__menu-icon--send-back::after\s*\{[\s\S]*?width: 7px;[\s\S]*?height: 7px;[\s\S]*?border: 1px solid #0b5a2f;[\s\S]*?background: #107c41;[\s\S]*?content: "";/,
    );
    expect(menusCss).not.toContain('content: "1"');
  });

  it('renders conditional-formatting symbol icons as vector marks, not text glyphs', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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
      expect(icon?.querySelector('path[fill="#107c41"]')).toBeTruthy();
    }
  });

  it('keeps Underline dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createUnderlineMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-underline-action]'));

    expect(menu.id).toBe('menu-underline');
    expect(items.map((item) => item.dataset.underlineAction)).toEqual(['single', 'double']);
    expect(items.map((item) => item.textContent)).toEqual(['下線', '二重下線']);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(2);

    expect(menusCss).toMatch(/#menu-underline\s*\{[\s\S]*?min-width: 118px;/);
    expect(menusCss).toMatch(
      /#menu-underline \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-underline \.fc-tb__menu-icon,[\s\S]*?#menu-underline \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Copy dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createCopyMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-copy-action]'));

    expect(menu.id).toBe('menu-copy');
    expect(items.map((item) => item.dataset.copyAction)).toEqual(['copy', 'picture']);
    expect(items.map((item) => item.textContent)).toEqual(['コピー', '図としてコピー...']);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(2);

    expect(menusCss).toMatch(/#menu-copy\s*\{[\s\S]*?min-width: 144px;/);
    expect(menusCss).toMatch(
      /#menu-copy \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-copy \.fc-tb__menu-icon,[\s\S]*?#menu-copy \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Paste dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

    expect(menusCss).toMatch(/#menu-paste\s*\{[\s\S]*?min-width: 198px;/);
    expect(menusCss).toMatch(
      /#menu-paste \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-paste \.fc-tb__menu-icon,[\s\S]*?#menu-paste \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Clear dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createClearMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-clear]'));

    expect(menu.id).toBe('menu-clear');
    expect(items.map((item) => item.dataset.clear)).toEqual([
      'all',
      'formats',
      'contents',
      'comments',
      'hyperlinks',
      'remove-hyperlinks',
      'conditional',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      'すべてクリア',
      '書式のクリア',
      '数式と値のクリア',
      'コメントとメモのクリア',
      'ハイパーリンクのクリア',
      'ハイパーリンクの削除',
      '条件付き書式のクリア',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(7);
    expect(menu.querySelector('path[fill="#f7e1ff"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#c00000"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-clear\s*\{[\s\S]*?min-width: 194px;/);
    expect(menusCss).toMatch(
      /#menu-clear \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-clear \.fc-tb__menu-icon,[\s\S]*?#menu-clear \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Sort and Filter dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createSortMenu('sortFilterHome');
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-sort]'));

    expect(menu.id).toBe('menu-sort-home');
    expect(items.map((item) => item.dataset.sort)).toEqual([
      'asc',
      'desc',
      'custom',
      'filter',
      'filter-by-value',
      'filter-clear',
      'filter-reapply',
      'filter-advanced',
      'dedupe',
      'conditional',
      'named',
    ]);
    expect(items.slice(0, 8).map((item) => item.textContent)).toEqual([
      '昇順で並べ替え',
      '降順で並べ替え',
      'ユーザー設定の並べ替え...',
      'フィルター',
      '選択したセルの値でフィルター',
      'クリア',
      '再適用',
      '詳細設定...',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(11);
    expect(menu.querySelector('path[stroke="#c00000"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-sort-home,[\s\S]*?#menu-sort\s*\{[\s\S]*?min-width: 218px;/);
    expect(menusCss).toMatch(
      /#menu-sort-home \.fc-tb__menu-item,[\s\S]*?#menu-sort \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-sort-home \.fc-tb__menu-icon,[\s\S]*?#menu-sort \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Find and Select dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createFindSelectMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-find-select]'));

    expect(menu.id).toBe('menu-find-select');
    expect(items.map((item) => item.dataset.findSelect)).toEqual([
      'find',
      'replace',
      'go-to',
      'go-to-special',
      'formulas',
      'comments',
      'conditional-format',
      'constants',
      'data-validation',
      'object-select',
      'selection-pane',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '検索...',
      '置換...',
      'ジャンプ...',
      '条件を選択してジャンプ...',
      '数式',
      'コメントとメモ',
      '条件付き書式',
      '定数',
      'データの入力規則',
      'オブジェクトの選択',
      '選択ウィンドウ...',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(11);
    expect(menu.querySelector('path[stroke="#8764b8"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(menu.querySelector('path[fill="#fff8cc"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-find-select\s*\{[\s\S]*?min-width: 182px;/);
    expect(menusCss).toMatch(
      /#menu-find-select \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-find-select \.fc-tb__menu-icon,[\s\S]*?#menu-find-select \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Fill dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createFillMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-fill]'));

    expect(menu.id).toBe('menu-fill');
    expect(items.map((item) => item.dataset.fill)).toEqual([
      'down',
      'right',
      'up',
      'left',
      'group',
      'series',
      'justify',
      'flash',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '下方向へコピー',
      '右方向へコピー',
      '上方向へコピー',
      '左方向へコピー',
      '作業グループへコピー...',
      '連続データの作成...',
      '文字の割付',
      'フラッシュ フィル',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(8);
    expect(menu.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    expect(menu.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(menu.querySelector('path[fill="#ed7d31"]')).toBeTruthy();

    expect(menusCss).toMatch(/#menu-fill\s*\{[\s\S]*?min-width: 178px;/);
    expect(menusCss).toMatch(
      /#menu-fill \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-fill \.fc-tb__menu-icon,[\s\S]*?#menu-fill \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
    for (const selector of [
      '.fc-tb__menu-icon--fill-days::after',
      '.fc-tb__menu-icon--fill-weekdays::after',
      '.fc-tb__menu-icon--fill-months::after',
      '.fc-tb__menu-icon--fill-years::after',
    ]) {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(menusCss).toMatch(
        new RegExp(`${escaped}[\\s\\S]*?border: 1px solid #107c41;[\\s\\S]*?content: "";`),
      );
    }
    const fillDateCss = menusCss.slice(
      menusCss.indexOf('.fc-tb__menu-icon--fill-days::before'),
      menusCss.indexOf('.fc-tb__menu-icon--freeze-col::before'),
    );
    for (const glyph of ['"D"', '"W"', '"M"', '"Y"']) {
      expect(fillDateCss).not.toContain(`content: ${glyph}`);
    }
  });

  it('keeps AutoSum dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const factories = createFormulasMenuFactories(ja.ribbonMenu, 'ja');
    const homeMenu = factories.createAutoSumMenu('autosum');
    const formulasMenu = factories.createAutoSumMenu('autosumFormula');
    const items = Array.from(homeMenu.querySelectorAll<HTMLButtonElement>('[data-autosum-fn]'));

    expect(homeMenu.id).toBe('menu-autosum-home');
    expect(formulasMenu.id).toBe('menu-autosum-formulas');
    expect(items.map((item) => item.dataset.autosumFn)).toEqual([
      'SUM',
      'AVERAGE',
      'COUNT',
      'MAX',
      'MIN',
      'MORE',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      '合計',
      '平均',
      '数値の個数',
      '最大値',
      '最小値',
      'その他の関数...',
    ]);
    expect(homeMenu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(
      1,
    );
    expect(homeMenu.querySelector('[data-autosum-fn="SUM"] .fc-tb__menu-icon--svg')).toBeTruthy();
    expect(homeMenu.querySelectorAll('.fc-tb__menu-item__icon-spacer')).toHaveLength(5);

    expect(menusCss).toMatch(
      /#menu-autosum-home,[\s\S]*?#menu-autosum-formulas\s*\{[\s\S]*?min-width: 128px;/,
    );
    expect(menusCss).toMatch(
      /#menu-autosum-home \.fc-tb__menu-item,[\s\S]*?#menu-autosum-formulas \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).not.toContain('.fc-tb__menu-icon--autosum-average::before');
    expect(menusCss).not.toContain('.fc-tb__menu-icon--autosum-more::after');
  });

  it('keeps Currency dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createStylesMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
    }).createCurrencyMenu();
    const presetItems = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('[data-currency-preset]'),
    );
    const footer = menu.querySelector<HTMLButtonElement>('[data-currency-footer]');

    expect(menu.id).toBe('menu-currency-home');
    expect(presetItems.map((item) => item.dataset.currencyPreset)).toEqual([
      '¥',
      '$',
      '€',
      '£',
      'CHF',
    ]);
    expect(presetItems.map((item) => item.textContent)).toEqual([
      '¥ 日本語',
      '$ 英語 (米国)',
      '€ ユーロ (€ 123)',
      '£ 英語 (英国)',
      'CHF フランス語 (スイス)',
    ]);
    expect(footer?.textContent).toBe('その他の通貨表示形式…');
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(0);
    expect(menu.querySelectorAll('.fc-tb__menu-item__icon-spacer')).toHaveLength(6);

    expect(menusCss).toMatch(/\.fc-tb__currency-menu\s*\{[\s\S]*?min-width: 190px;/);
    expect(menusCss).toMatch(
      /\.fc-tb__currency-menu \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__currency-menu \.fc-tb__menu-item__icon-spacer\s*\{[\s\S]*?width: 0;/,
    );
    expect(menusCss).not.toContain('.fc-tb__menu-icon--currency-yen::before');
    expect(menusCss).not.toContain('.fc-tb__menu-icon--currency-more::before');
  });

  it('keeps Cell Styles gallery geometry close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createStylesMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
    }).createCellStylesMenu();
    const scrollBody = menu.querySelector<HTMLElement>(':scope > .fc-tb__cellstyle-scroll');
    const headings = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__cellstyle-heading') ?? [],
    ).map((heading) => heading.textContent);
    const grids = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__cellstyle-grid') ?? [],
    );

    expect(menu.id).toBe('menu-cell-styles-home');
    expect(menu.classList.contains('fc-tb__cellstyle-menu')).toBe(true);
    expect(headings).toEqual([
      '良い、悪い、標準',
      'データとモデル',
      'タイトルと見出し',
      'テーマのセル スタイル',
      '表示形式',
    ]);
    expect(grids.length).toBeGreaterThanOrEqual(5);
    expect(menu.querySelector('[data-cell-style="normal"]')?.textContent).toBe('標準');
    expect(menu.querySelector('[data-cell-style="good"]')?.textContent).toBe('良い');
    expect(menu.querySelectorAll('.fc-tb__cellstyle-footer')).toHaveLength(2);

    expect(menusCss).toMatch(/\.fc-tb__cellstyle-menu\s*\{[\s\S]*?width: 566px;/);
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-scroll\s*\{[\s\S]*?max-height: min\(388px, calc\(100vh - 244px\)\);/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-grid\s*\{[\s\S]*?grid-template-columns: repeat\(6, 82px\);[\s\S]*?gap: 7px 11px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-heading\s*\{[\s\S]*?background: #f3f2f1;[\s\S]*?font-weight: 400;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__cellstyle-chip\s*\{[\s\S]*?min-width: 82px;[\s\S]*?height: 22px;[\s\S]*?border-radius: 0;/,
    );
  });

  it('keeps Format as Table gallery geometry close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createStylesMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
    }).createTableStyleMenu('formatTableHome');
    const scrollBody = menu.querySelector<HTMLElement>(':scope > .fc-tb__tablestyle-scroll');
    const headings = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__tablestyle-heading') ?? [],
    ).map((heading) => heading.textContent);
    const grids = Array.from(
      scrollBody?.querySelectorAll<HTMLElement>(':scope > .fc-tb__tablestyle-grid') ?? [],
    );

    expect(menu.id).toBe('menu-table-style-home');
    expect(menu.classList.contains('fc-tb__tablestyle-menu')).toBe(true);
    expect(headings).toEqual(['淡色', '中間', '濃色']);
    expect(grids.map((grid) => grid.querySelectorAll('.fc-tb__tablestyle-swatch').length)).toEqual([
      28, 28, 7,
    ]);
    expect(menu.querySelectorAll('.fc-tb__tablestyle-footer')).toHaveLength(2);

    expect(menusCss).toMatch(/\.fc-tb__tablestyle-menu\s*\{[\s\S]*?width: 515px;/);
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-scroll\s*\{[\s\S]*?max-height: min\(496px, calc\(100vh - 116px\)\);/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-grid\s*\{[\s\S]*?grid-template-columns: repeat\(7, 62px\);[\s\S]*?gap: 8px 9px;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-heading\s*\{[\s\S]*?background: #f3f2f1;[\s\S]*?font-weight: 400;/,
    );
    expect(menusCss).toMatch(
      /\.fc-tb__tablestyle-swatch\s*\{[\s\S]*?width: 62px;[\s\S]*?height: 47px;[\s\S]*?border-radius: 0;/,
    );
  });

  it('keeps Insert and Delete Cells dropdowns close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const factories = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    });
    const insertMenu = factories.createInsertCellsMenu();
    const deleteMenu = factories.createDeleteCellsMenu();
    const insertItems = Array.from(
      insertMenu.querySelectorAll<HTMLButtonElement>('[data-cell-insert]'),
    );
    const deleteItems = Array.from(
      deleteMenu.querySelectorAll<HTMLButtonElement>('[data-cell-delete]'),
    );

    expect(insertMenu.id).toBe('menu-insert-cells');
    expect(insertItems.map((item) => item.dataset.cellInsert)).toEqual([
      'cells',
      'rows',
      'cols',
      'sheet',
    ]);
    expect(insertItems.map((item) => item.textContent)).toEqual([
      'セルを挿入...',
      'シートの行を挿入',
      'シートの列を挿入',
      'シートの挿入',
    ]);
    expect(deleteMenu.id).toBe('menu-delete-cells');
    expect(deleteItems.map((item) => item.dataset.cellDelete)).toEqual([
      'cells',
      'rows',
      'cols',
      'row',
      'col',
      'sheet',
    ]);
    expect(deleteItems.map((item) => item.textContent)).toEqual([
      'セルを削除...',
      'シートの行を削除',
      'シートの列を削除',
      '行の削除',
      '列の削除',
      'シートの削除',
    ]);
    expect(
      insertMenu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg'),
    ).toHaveLength(4);
    expect(
      deleteMenu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg'),
    ).toHaveLength(6);
    expect(insertMenu.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(deleteMenu.querySelector('path[stroke="#c00000"]')).toBeTruthy();

    expect(menusCss).toMatch(
      /#menu-insert-cells,[\s\S]*?#menu-delete-cells\s*\{[\s\S]*?min-width: 166px;/,
    );
    expect(menusCss).toMatch(
      /#menu-insert-cells \.fc-tb__menu-item,[\s\S]*?#menu-delete-cells \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-insert-cells \.fc-tb__menu-icon,[\s\S]*?#menu-delete-cells \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Format Cells dropdown section chrome close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createFormatCellsMenu();
    const headings = Array.from(menu.children)
      .filter((child): child is HTMLElement => child instanceof HTMLElement)
      .filter((child) => child.classList.contains('fc-tb__menu-heading'))
      .map((heading) => heading.textContent);
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-cell-format]'));
    const submenuTriggers = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('[data-format-submenu]'),
    );
    const visibilityItems = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('#menu-format-cells-visibility [data-cell-format]'),
    );
    const tabColorMenu = menu.querySelector<HTMLElement>('#menu-format-cells-tabColor');
    const tabColorHeadings = Array.from(
      tabColorMenu?.querySelectorAll<HTMLElement>('.fc-tb__menu-heading') ?? [],
    ).map((heading) => heading.textContent);

    expect(menu.id).toBe('menu-format-cells');
    expect(headings).toEqual(['セルのサイズ', '表示設定', 'シートの整理', '保護']);
    expect(submenuTriggers.map((item) => item.dataset.formatSubmenu)).toEqual([
      'visibility',
      'tabColor',
    ]);
    expect(submenuTriggers.map((item) => item.textContent)).toEqual([
      '非表示/再表示',
      'シート見出しの色',
    ]);
    expect(menu.querySelector('#menu-format-cells-visibility')).toBeTruthy();
    expect(menu.querySelector('#menu-format-cells-tabColor')).toBeTruthy();
    expect(visibilityItems.map((item) => item.dataset.cellFormat)).toEqual([
      'hide-rows',
      'hide-cols',
      'hide-sheet',
      'show-rows',
      'show-cols',
      'unhide-sheet',
    ]);
    expect(visibilityItems.map((item) => item.textContent)).toEqual([
      '行を表示しない',
      '列を表示しない',
      'シートを表示しない',
      '行の再表示',
      '列の再表示',
      'シートの再表示...',
    ]);
    expect(menu.querySelector('[data-cell-format="rename-sheet"]')?.textContent).toBe(
      'シート名の変更',
    );
    expect(items.map((item) => item.dataset.cellFormat)).toContain('move-sheet-copy');
    expect(items.map((item) => item.dataset.cellFormat)).not.toContain('move-sheet-left');
    expect(items.map((item) => item.dataset.cellFormat)).not.toContain('move-sheet-right');
    expect(tabColorHeadings).toEqual(['テーマの色', '標準の色']);
    expect(tabColorMenu?.querySelector('[data-cell-format="tab-color-none"]')?.textContent).toBe(
      '色なし',
    );
    expect(tabColorMenu?.querySelector('[data-cell-format="tab-color-more"]')?.textContent).toBe(
      'その他の色…',
    );
    expect(tabColorMenu?.querySelectorAll('.fc-tb__color-swatch')).toHaveLength(14);
    expect(items.at(-1)?.dataset.cellFormat).toBe('dialog');
    expect(items.at(-1)?.textContent).toBe('セルの書式設定...');
    expect(items.map((item) => item.dataset.cellFormat)).toContain('lock-cell');
    expect(
      menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg').length,
    ).toBeGreaterThan(10);

    expect(menusCss).toMatch(/#menu-format-cells\s*\{[\s\S]*?min-width: 208px;/);
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__menu-heading\s*\{[\s\S]*?color: #808080;[\s\S]*?font-weight: 400;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__menu-item--checked \.fc-tb__menu-item__text::before\s*\{[\s\S]*?border-bottom: 2px solid #107c41;[\s\S]*?border-left: 2px solid #107c41;[\s\S]*?content: "";[\s\S]*?transform: rotate\(-45deg\);/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__submenu--format\s*\{[\s\S]*?min-width: 172px;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__submenu--format-tab-color\s*\{[\s\S]*?min-width: 206px;/,
    );
    expect(menusCss).toMatch(
      /#menu-format-cells \.fc-tb__submenu--format-tab-color \.fc-tb__color-swatch-grid\s*\{[\s\S]*?grid-template-columns: repeat\(7, 18px\);/,
    );
  });

  it('renders checked, reapply, and warning fallback menu glyphs as vector overlays', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');

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

  it('keeps Merge Cells dropdown compact and close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createMergeMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-merge-action]'));

    expect(menu.id).toBe('menu-merge');
    expect(items.map((item) => item.dataset.mergeAction)).toEqual([
      'mergeCenter',
      'mergeAcross',
      'mergeCells',
      'unmergeCells',
    ]);
    expect(items.map((item) => item.textContent)).toEqual([
      'セルを結合して中央揃え',
      '横方向に結合',
      'セルの結合',
      'セル結合の解除',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(4);

    expect(menusCss).toMatch(/#menu-merge\s*\{[\s\S]*?min-width: 206px;/);
    expect(menusCss).toMatch(
      /#menu-merge \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-merge \.fc-tb__menu-icon,[\s\S]*?#menu-merge \.fc-tb__menu-icon-svg\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('keeps Wrap Text dropdown close to Japanese Excel 365 desktop', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const ja = dictionaries.ja;
    const menu = createHomeMenuFactories({
      ribbonLang: 'ja',
      ribbonMenuText: ja.ribbonMenu,
      ribbonText: ja.ribbon,
      formatDialog: ja.formatDialog,
      sheetTabs: ja.sheetTabs,
      viewToolbar: ja.viewToolbar,
    }).createWrapMenu();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-wrap-action]'));

    expect(menu.id).toBe('menu-wrap');
    expect(items.map((item) => item.dataset.wrapAction)).toEqual(['wrapText', 'shrinkToFit']);
    expect(items.map((item) => item.textContent)).toEqual([
      '折り返して全体を表示',
      '縮小して全体を表示する',
    ]);
    expect(menu.querySelectorAll('.fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')).toHaveLength(1);
    expect(menu.querySelectorAll('.fc-tb__menu-item__icon-spacer')).toHaveLength(1);

    expect(menusCss).toMatch(/#menu-wrap\s*\{[\s\S]*?min-width: 220px;/);
    expect(menusCss).toMatch(
      /#menu-wrap \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 27px;[\s\S]*?padding: 3px 12px 3px 20px;/,
    );
    expect(menusCss).toMatch(
      /#menu-wrap \.fc-tb__menu-icon,[\s\S]*?#menu-wrap \.fc-tb__menu-icon-svg,[\s\S]*?#menu-wrap \.fc-tb__menu-item__icon-spacer\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
  });

  it('adds stable modifier classes to Conditional Formatting icon-set choices', () => {
    const menu = createConditionalMenu('ja');
    const arrows = menu.querySelector<HTMLElement>('[data-cf-action="icons-arrows5"]');
    const traffic = menu.querySelector<HTMLElement>('[data-cf-action="icons-traffic3"]');
    const flags = menu.querySelector<HTMLElement>('[data-cf-action="icons-flags3"]');
    const symbols = menu.querySelector<HTMLElement>('[data-cf-action="icons-symbols3"]');

    expect(arrows?.classList.contains('fc-tb__cf-icon-choice--icons-arrows5')).toBe(true);
    expect(traffic?.classList.contains('fc-tb__cf-icon-choice--icons-traffic3')).toBe(true);
    expect(flags?.classList.contains('fc-tb__cf-icon-choice--icons-flags3')).toBe(true);
    expect(arrows?.querySelectorAll('span')).toHaveLength(5);
    expect(traffic?.querySelectorAll('span')).toHaveLength(3);
    expect(symbols?.querySelectorAll('span')).toHaveLength(3);
    expect(arrows?.textContent).toBe('');
    expect(symbols?.textContent).toBe('');

    const conditionalSource = readFileSync(
      join(root, 'src/toolbar/ribbon/menus/conditional.ts'),
      'utf8',
    );
    for (const glyph of ['▲', '↗', '▶', '↘', '▼', '★', '✓', '×', '⚑', '◔', '▮', '■']) {
      expect(conditionalSource).not.toContain(glyph);
    }
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

  it('shows SVG previews for Borders dropdown footer and drawing tools', () => {
    const menu = createBordersMenu({
      ribbonText: {
        bottomBorder: 'Bottom Border',
        topBorder: 'Top Border',
        leftBorder: 'Left Border',
        rightBorder: 'Right Border',
        noBorder: 'No Border',
        allBorders: 'All Borders',
        outsideBorders: 'Outside Borders',
        thickOutsideBorders: 'Thick Outside Borders',
        doubleBottomBorder: 'Double Bottom Border',
        thickBottomBorder: 'Thick Bottom Border',
        topAndBottomBorder: 'Top and Bottom Border',
        topAndThickBottomBorder: 'Top and Thick Bottom Border',
        topAndDoubleBottomBorder: 'Top and Double Bottom Border',
        drawBordersHeading: 'Draw Borders',
        drawBorder: 'Draw Border',
        drawBorderGrid: 'Draw Border Grid',
        eraseBorder: 'Erase Border',
        lineColor: 'Line Color',
        lineStyle: 'Line Style',
        lineStyleNone: 'None',
        moreBorders: 'More Borders...',
        themeColors: 'Theme Colors',
        standardColors: 'Standard Colors',
        automatic: 'Automatic',
      } as Parameters<typeof createBordersMenu>[0]['ribbonText'],
      getBorderColor: () => '#000000',
      onPickColor: vi.fn(),
    });

    expect(menu.querySelector('[data-border-preset="format"] .fc-tb__border-preview')).toBeTruthy();
    expect(
      menu.querySelector('[data-border-draw="erase"] .fc-tb__border-preview--eraser'),
    ).toBeTruthy();
    expect(
      menu.querySelector('[data-border-submenu="lineColor"] .fc-tb__border-preview--line-color'),
    ).toBeTruthy();
    expect(
      menu.querySelector('[data-border-submenu="lineStyle"] .fc-tb__border-preview--line-style'),
    ).toBeTruthy();
  });

  it('keeps Borders dropdown close to Japanese Excel 365 desktop menu structure', () => {
    const menuCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    const menu = createBordersMenu({
      ribbonText: {
        bottomBorder: '下罫線',
        topBorder: '上罫線',
        leftBorder: '左罫線',
        rightBorder: '右罫線',
        noBorder: '罫線なし',
        allBorders: '格子',
        outsideBorders: '外枠',
        thickOutsideBorders: '外枠太罫線',
        doubleBottomBorder: '下二重罫線',
        thickBottomBorder: '下太罫線',
        topAndBottomBorder: '上罫線 + 下罫線',
        topAndThickBottomBorder: '上罫線 + 下太罫線',
        topAndDoubleBottomBorder: '上罫線 + 下二重罫線',
        drawBordersHeading: '罫線の作成',
        drawBorder: '罫線の作成',
        drawBorderGrid: '罫線グリッドの作成',
        eraseBorder: '罫線の削除',
        lineColor: '線の色',
        lineStyle: '線のスタイル',
        lineStyleNone: 'なし',
        moreBorders: 'その他の罫線...',
        themeColors: 'テーマの色',
        standardColors: '標準の色',
        automatic: '自動',
      } as Parameters<typeof createBordersMenu>[0]['ribbonText'],
      getBorderColor: () => '#000000',
      onPickColor: vi.fn(),
    });

    expect(menu.classList.contains('fc-tb__menu--borders')).toBe(true);
    expect(menu.querySelectorAll('[role="separator"]')).toHaveLength(3);
    expect(menu.querySelector('.fc-tb__menu-heading')?.textContent).toBe('罫線の作成');
    expect(menu.querySelectorAll('[data-border-preset]')).toHaveLength(14);
    expect(menu.querySelectorAll('[data-border-draw]')).toHaveLength(3);
    expect(menu.querySelectorAll('[data-border-submenu]')).toHaveLength(2);
    expect(
      menu.querySelector('.fc-tb__submenu--line-style .fc-tb__submenu-item--line-style-none')
        ?.textContent,
    ).toBe('なし');
    expect(menu.querySelectorAll('.fc-tb__submenu--line-style .fc-tb__line-sample')).toHaveLength(
      11,
    );

    expect(menuCss).toMatch(/\.fc-tb__menu--borders\s*\{[\s\S]*?min-width: 186px;/);
    expect(menuCss).toMatch(
      /\.fc-tb__menu--borders \.fc-tb__menu-item\s*\{[\s\S]*?min-height: 24px;[\s\S]*?padding: 2px 12px 2px 18px;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__menu--borders \.fc-tb__menu-item\[aria-expanded="true"\]\s*\{[\s\S]*?background: #107c41;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__menu--borders \.fc-tb__submenu--line-style\s*\{[\s\S]*?min-width: 101px;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__submenu--line-style \.fc-tb__line-sample\s*\{[\s\S]*?width: 74px;/,
    );
    expect(menuCss).toMatch(
      /\.fc-tb__submenu--line-color \.fc-colorpalette__action--automatic\s*\{[\s\S]*?border-color: #107c41;/,
    );
  });

  it('keeps raw menu button creation centralized in createMenuButton', () => {
    const directButtons = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes("document.createElement('button')"))
      .map(({ name }) => name);

    expect(directButtons).toEqual([]);
  });

  it('keeps Paste menu labels backed by shared i18n dictionaries', () => {
    const pasteSource = readFileSync(join(menusDir, 'paste.ts'), 'utf8');
    const toolbarDefaultsSource = readFileSync(join(mountDir, 'toolbar-defaults.ts'), 'utf8');

    expect(pasteSource).toContain('import type { Strings }');
    expect(pasteSource).toContain('menuIconButton(t.ribbon.paste,');
    expect(pasteSource).toContain('menuIconButton(pasteText.pasteFormulas,');
    expect(pasteSource).toContain('menuIconButton(pasteText.pasteValues,');
    expect(pasteSource).toContain('menuIconButton(pasteText.pasteSpecialDialog,');
    expect(pasteSource).not.toContain('const ja =');
    expect(pasteSource).not.toContain('貼り付け');
    expect(pasteSource).not.toContain('Paste Special');
    expect(toolbarDefaultsSource).toContain('createPasteMenu(dictionaries[lang])');
  });

  it('keeps Home Insert and Delete cell menu labels backed by ribbonMenu strings', () => {
    const homeSource = readFileSync(join(menusDir, 'home.ts'), 'utf8');

    expect(homeSource).toContain("menuIconButton(t.insertCells, 'cellInsert', 'cells'");
    expect(homeSource).toContain("menuIconButton(t.insertRows, 'cellInsert', 'rows'");
    expect(homeSource).toContain("menuIconButton(t.insertCols, 'cellInsert', 'cols'");
    expect(homeSource).toContain("menuIconButton(t.deleteCells, 'cellDelete', 'cells'");
    expect(homeSource).toContain("menuIconButton(t.deleteRows, 'cellDelete', 'rows'");
    expect(homeSource).toContain("menuIconButton(t.deleteCols, 'cellDelete', 'cols'");
    expect(homeSource).not.toContain('セルを挿入');
    expect(homeSource).not.toContain('Insert Cells');
    expect(homeSource).not.toContain('シートの行を挿入');
    expect(homeSource).not.toContain('Delete Sheet Rows');
  });

  it('keeps Underline split menu labels backed by ribbonMenu strings', () => {
    const homeSource = readFileSync(join(menusDir, 'home.ts'), 'utf8');

    expect(homeSource).toContain("menuIconButton(t.underlineSingle, 'underlineAction'");
    expect(homeSource).toContain('t.underlineDouble');
    expect(homeSource).not.toContain('二重下線');
    expect(homeSource).not.toContain('Double Underline');
    expect(homeSource).not.toContain('const ja =');
  });

  it('keeps Calculation Options menu labels backed by ribbonMenu strings', () => {
    const formulasSource = readFileSync(join(menusDir, 'formulas.ts'), 'utf8');

    expect(formulasSource).toContain('calcOptionButton(t.calcAutomatic,');
    expect(formulasSource).toContain('calcOptionButton(t.calcAutoNoTable,');
    expect(formulasSource).toContain('calcOptionButton(t.calcManual,');
    expect(formulasSource).toContain('calcOptionButton(t.calcCalculateNow,');
    expect(formulasSource).toContain('calcOptionButton(t.calcCalculateSheet,');
    expect(formulasSource).toContain('calcOptionButton(t.calcIterative,');
    expect(formulasSource).not.toContain('const ja =');
    expect(formulasSource).not.toContain('Calculate Now');
    expect(formulasSource).not.toContain('再計算実行');
  });

  it('keeps Table and Cell style custom section labels backed by ribbonMenu strings', () => {
    const stylesSource = readFileSync(join(menusDir, 'styles.ts'), 'utf8');

    expect(stylesSource).toContain('t.tableStyleCustom');
    expect(stylesSource).toContain('t.pivotTableStyleCustom');
    expect(stylesSource).toContain('t.cellStyleCustom');
    expect(stylesSource).not.toContain("ribbonLang === 'ja' ? 'ユーザー設定'");
    expect(stylesSource).not.toContain('Custom PivotTable');
  });

  it('keeps Table and Cell style dialogs backed by required ribbonMenu strings', () => {
    const tableStyleDialogSource = readFileSync(
      join(root, 'src/toolbar/dialogs/table-style.ts'),
      'utf8',
    );
    const cellStyleDialogSource = readFileSync(
      join(root, 'src/toolbar/dialogs/cell-style.ts'),
      'utf8',
    );

    expect(tableStyleDialogSource).toContain('tableStyleName: string');
    expect(tableStyleDialogSource).toContain('t.tableStyleName');
    expect(tableStyleDialogSource).toContain('t.tableStyleMedium');
    expect(tableStyleDialogSource).toContain('t.tableStyleBandedRows');
    expect(cellStyleDialogSource).toContain('cellStyleName: string');
    expect(cellStyleDialogSource).toContain('t.cellStyleName');
    expect(cellStyleDialogSource).toContain('t.cellStyleNormal');
    expect(cellStyleDialogSource).toContain('t.cellStyleIncludeProtection');
    for (const source of [tableStyleDialogSource, cellStyleDialogSource]) {
      expect(source).not.toContain('menuText(');
      expect(source).not.toContain("'Style name'");
      expect(source).not.toContain("'Medium'");
      expect(source).not.toContain("'Normal'");
      expect(source).not.toContain("'Style includes'");
      expect(source).not.toContain("'First column emphasis'");
    }
  });

  it('keeps Cell Styles merge report text backed by required ribbonMenu strings', () => {
    const defaultsSource = readFileSync(join(mountDir, 'dynamic-dropdowns-defaults.ts'), 'utf8');

    expect(defaultsSource).toContain('strings.ribbonMenu.cellStyleMergeImported.replace');
    expect(defaultsSource).not.toContain('cellStyleMergeImported?:');
    expect(defaultsSource).not.toContain('style(s) imported');
  });

  it('keeps Create Table dialog labels backed by shared dialog strings', () => {
    const defaultsSource = readFileSync(join(mountDir, 'dynamic-dropdowns-defaults.ts'), 'utf8');

    expect(defaultsSource).toContain('pivotDialogStrings.createTableTitle');
    expect(defaultsSource).toContain('pivotDialogStrings.createTableRangeLabel');
    expect(defaultsSource).toContain('pivotDialogStrings.createTableHeadersLabel');
    expect(defaultsSource).toContain('pivotDialogStrings.createTableInvalidRange');
    expect(defaultsSource).not.toContain('テーブルの作成');
    expect(defaultsSource).not.toContain('Create Table');
    expect(defaultsSource).not.toContain('My table has headers');
  });

  it('keeps Fill Series dialog labels backed by shared dialog strings', () => {
    const fillSeriesSource = readFileSync(join(ribbonDir, 'fill-series.ts'), 'utf8');
    const defaultsSource = readFileSync(join(mountDir, 'dynamic-dropdowns-defaults.ts'), 'utf8');

    expect(fillSeriesSource).toContain("Strings['fillSeriesDialog']");
    expect(fillSeriesSource).toContain('createDialogShell({ title })');
    expect(fillSeriesSource).toContain('appendDialogActions(shell.footer');
    expect(fillSeriesSource).toContain('installDialogLifecycle<');
    expect(fillSeriesSource).toContain('mountDialog(shell');
    expect(fillSeriesSource).toContain('t.seriesIn');
    expect(fillSeriesSource).toContain('t.autoFill');
    expect(fillSeriesSource).toContain('t.weekday');
    expect(defaultsSource).toContain('fillSeriesDialogStrings');
    expect(defaultsSource).toContain('fillSeriesDialog: {');
    expect(fillSeriesSource).not.toContain("ribbonLang === 'ja'");
    expect(fillSeriesSource).not.toContain('const ja =');
    expect(fillSeriesSource).not.toContain('連続データ');
    expect(fillSeriesSource).not.toContain('AutoFill');
    expect(fillSeriesSource).not.toContain("'Cancel'");
    expect(fillSeriesSource).not.toContain('"Cancel"');
    expect(fillSeriesSource).not.toContain("const overlay = document.createElement('div')");
    expect(fillSeriesSource).not.toContain("const cancelBtn = document.createElement('button')");
    expect(fillSeriesSource).not.toContain("const okBtn = document.createElement('button')");
  });

  it('keeps Home Format action prompt labels backed by ribbonMenu strings', () => {
    const cellFormatSource = readFileSync(join(ribbonDir, 'cell-format-action.ts'), 'utf8');
    const dynamicDefaultsSource = readFileSync(
      join(mountDir, 'dynamic-dropdowns-defaults.ts'),
      'utf8',
    );

    expect(cellFormatSource).toContain('type CellFormatMenuText');
    expect(cellFormatSource).toContain('showRenameSheetDialog');
    expect(cellFormatSource).toContain('t.sheetNameLabel');
    expect(cellFormatSource).toContain('t.sheetNameRequired');
    expect(cellFormatSource).toContain('t.rowHeightLabel');
    expect(cellFormatSource).toContain('t.colWidthLabel');
    expect(dynamicDefaultsSource).toContain('showDimensionDialog({');
    expect(cellFormatSource).not.toContain("ribbonLang === 'ja' ? 'シート名'");
    expect(cellFormatSource).not.toContain('Enter a sheet name.');
    expect(cellFormatSource).not.toContain('Height (px)');
    expect(cellFormatSource).not.toContain('Width (px)');
  });

  it('keeps toolbar dialog select option creation centralized in form-controls', () => {
    const sortSource = readFileSync(join(root, 'src/toolbar/dialogs/sort.ts'), 'utf8');
    const conditionalFormatSource = readFileSync(
      join(root, 'src/toolbar/dialogs/conditional-format.ts'),
      'utf8',
    );
    const scriptCommandSource = readFileSync(
      join(root, 'src/toolbar/dialogs/script-command.ts'),
      'utf8',
    );
    const formatDialogTabSources = [
      'src/interact/format-dialog-tabs/align.ts',
      'src/interact/format-dialog-tabs/fill.ts',
      'src/interact/format-dialog-tabs/border.ts',
      'src/interact/format-dialog-tabs/font.ts',
      'src/interact/format-dialog-view.ts',
      'src/interact/format-dialog.ts',
      'src/interact/format-dialog-tabs/more.ts',
    ].map((path) => readFileSync(join(root, path), 'utf8'));
    const interactSurfaceSources = [
      'src/interact/filter-dropdown.ts',
      'src/interact/view-toolbar.ts',
      'src/interact/pivot-field-settings.ts',
      'src/interact/pivot-table-dialog.ts',
      'src/interact/workbook-objects.ts',
      'src/interact/page-setup-dialog.ts',
      'src/interact/conditional-dialog.ts',
      'src/interact/cf-rules-dialog.ts',
      'src/interact/named-range-dialog.ts',
      'src/interact/find-replace.ts',
      'src/interact/fx-dialog.ts',
    ].map((path) => readFileSync(join(root, path), 'utf8'));

    for (const source of [
      sortSource,
      conditionalFormatSource,
      scriptCommandSource,
      ...formatDialogTabSources,
      ...interactSurfaceSources,
    ]) {
      expect(source).toMatch(
        /createDialogSelect|appendDialogSelectOptions|appendDialogDatalistOptions/,
      );
      expect(source.match(/document\.createElement\('option'\)/g) ?? []).toHaveLength(0);
      expect(source.match(/appendChild\(option\)/g) ?? []).toHaveLength(0);
      expect(source.match(/appendChild\(opt\)/g) ?? []).toHaveLength(0);
    }
  });

  it('keeps Conditional Formatting date choice labels backed by conditionalMenu strings', () => {
    const conditionalMenuSource = readFileSync(join(menusDir, 'conditional.ts'), 'utf8');
    const actionSource = readFileSync(join(ribbonDir, 'conditional-menu-action.ts'), 'utf8');

    expect(conditionalMenuSource).toContain('datePeriods: {');
    expect(conditionalMenuSource).toContain('dateUnsupported: t.dateUnsupported');
    expect(conditionalMenuSource).toContain('ok: t.ok');
    expect(conditionalMenuSource).toContain('cancel: t.cancel');
    expect(actionSource).toContain('cfDatePeriodOptions(title.datePeriods)');
    expect(actionSource).toContain('okLabel: title.ok');
    expect(actionSource).toContain('cancelLabel: title.cancel');
    expect(actionSource).toContain('message: title.dateUnsupported');
    expect(actionSource).not.toContain("ribbonLang === 'ja'");
    expect(actionSource).not.toContain('昨日');
    expect(actionSource).not.toContain('Yesterday');
    expect(actionSource).not.toContain('Enter one of the supported date conditions.');
  });

  it('keeps select/dropdown chrome labels backed by ribbon strings', () => {
    const selectColorSource = readFileSync(join(ribbonDir, 'select-color.ts'), 'utf8');
    const buttonSource = readFileSync(join(ribbonDir, 'button.ts'), 'utf8');
    const dropdownCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/dropdowns.css'), 'utf8');

    expect(selectColorSource).toContain("import { createRibbonButton } from './button.js'");
    expect(selectColorSource).toContain('const createRibbonControlButton');
    expect(selectColorSource).toContain('createRibbonControlButton({');
    expect(selectColorSource).not.toContain("document.createElement('button')");
    expect(buttonSource).toContain("document.createElement('button')");
    expect(selectColorSource).not.toContain("const item = document.createElement('button')");
    expect(selectColorSource).toContain('ribbonText.fontSectionTheme');
    expect(selectColorSource).toContain('ribbonText.fontSectionRecent');
    expect(selectColorSource).toContain('ribbonText.fontSectionAll');
    expect(selectColorSource).toContain('ribbonText.fontRoleHeading');
    expect(selectColorSource).toContain('ribbonText.fontRoleBody');
    expect(selectColorSource).toContain('ribbonText.currentView');
    expect(selectColorSource).toContain('ribbonText.marginsCustomDialog');
    expect(selectColorSource).toContain('ribbonText.marginTop');
    expect(selectColorSource).not.toContain("arrow.textContent = '›'");
    expect(dropdownCss).toMatch(
      /\.fc-tb__rb-dd__submenu\s*\{[\s\S]*?border-top: 4px solid transparent;[\s\S]*?border-bottom: 4px solid transparent;[\s\S]*?border-left: 5px solid var\(--fc-tb-fg\);/,
    );
    expect(selectColorSource).not.toContain('テーマのフォント');
    expect(selectColorSource).not.toContain('Theme Fonts');
    expect(selectColorSource).not.toContain('Current view');
    expect(selectColorSource).not.toContain('Custom margins...');
  });

  it('keeps control dispatch defaults and prompt labels backed by shared strings', () => {
    const controlDispatchSource = readFileSync(join(ribbonDir, 'control-dispatch.ts'), 'utf8');

    expect(controlDispatchSource).toContain('ribbonText.defaultFontFamily');
    expect(controlDispatchSource).toContain('ribbonText.defaultFontSize');
    expect(controlDispatchSource).toContain('showPageScaleDialog');
    expect(controlDispatchSource).toContain('okLabel: pageScaleText.ok');
    expect(controlDispatchSource).toContain('cancelLabel: pageScaleText.cancel');
    expect(controlDispatchSource).not.toContain('showNumberPrompt');
    expect(controlDispatchSource).not.toContain('showPrompt');
    expect(controlDispatchSource).not.toContain("ribbonLang === 'ja' ? '游ゴシック Regular'");
    expect(controlDispatchSource).not.toContain("ribbonLang === 'ja' ? 12");
    expect(controlDispatchSource).not.toContain("okLabel: 'OK'");
    expect(controlDispatchSource).not.toContain("ribbonLang === 'ja' ? 'キャンセル'");
  });

  it('keeps backstage title search status backed by shell strings', () => {
    const backstageTitleSource = readFileSync(join(ribbonDir, 'backstage-title.ts'), 'utf8');

    expect(backstageTitleSource).toContain('findNoMatches: string');
    expect(backstageTitleSource).toContain("shellText.findNoMatches.replace('{query}', query)");
    expect(backstageTitleSource).not.toContain(`ribbonLang === 'ja' ? \`「\${query}」`);
    expect(backstageTitleSource).not.toContain(`No matches for "\${query}"`);
  });

  it('uses shared localized labels for report dialogs from default toolbar glue', () => {
    const dynamicDefaultsSource = readFileSync(
      join(mountDir, 'dynamic-dropdowns-defaults.ts'),
      'utf8',
    );
    const toolbarDefaultsSource = readFileSync(join(mountDir, 'toolbar-defaults.ts'), 'utf8');
    const reportSource = readFileSync(join(root, 'src/toolbar/dialogs/report.ts'), 'utf8');
    const dialogsIndexSource = readFileSync(join(root, 'src/toolbar/dialogs/index.ts'), 'utf8');
    const indexSource = readFileSync(join(root, 'src/index.ts'), 'utf8');

    for (const source of [dynamicDefaultsSource, toolbarDefaultsSource]) {
      const calls = source.match(/showReport\(\{/g) ?? [];
      const sharedLabelSpreads = source.match(/\.\.\.reportDialogLabels\(/g) ?? [];
      expect(sharedLabelSpreads.length).toBe(calls.length);
      expect(source).not.toContain('emptyLabel: strings.reviewReports.noIssues');
      expect(source).not.toContain('closeLabel: strings.workbookObjects.close');
      expect(source).not.toContain('infoLabel: strings.reviewReports.info');
      expect(source).not.toContain('warningLabel: strings.reviewReports.warning');
    }
    expect(reportSource).toContain('export const reportDialogLabels');
    expect(reportSource).toContain('emptyLabel: strings.reviewReports.noIssues');
    expect(reportSource).toContain('closeLabel: strings.workbookObjects.close');
    expect(reportSource).toContain('infoLabel: strings.reviewReports.info');
    expect(reportSource).toContain('warningLabel: strings.reviewReports.warning');
    expect(reportSource).toContain('emptyLabel: string');
    expect(reportSource).toContain('closeLabel: string');
    expect(reportSource).toContain('infoLabel: string');
    expect(reportSource).toContain('warningLabel: string');
    expect(dialogsIndexSource).toContain('reportDialogLabels');
    expect(dialogsIndexSource).toContain('type ReportDialogLabels');
    expect(indexSource).toContain('reportDialogLabels');
    expect(indexSource).toContain('ReportDialogLabels');
    expect(reportSource).not.toContain('No issues found.');
    expect(reportSource).not.toContain("'Warning'");
    expect(reportSource).not.toContain("'Info'");
    expect(reportSource).not.toContain("'Close'");
  });

  it('keeps default dialog prompt labels backed by shared strings', () => {
    const controlDispatchSource = readFileSync(join(ribbonDir, 'control-dispatch.ts'), 'utf8');
    const dynamicDefaultsSource = readFileSync(
      join(mountDir, 'dynamic-dropdowns-defaults.ts'),
      'utf8',
    );
    const toolbarDefaultsSource = readFileSync(join(mountDir, 'toolbar-defaults.ts'), 'utf8');
    const dialogSources = [
      'advanced-filter.ts',
      'format-as-table.ts',
      'page-scale.ts',
      'rename-sheet.ts',
      'remove-duplicates.ts',
      'sort.ts',
      'zoom.ts',
    ].map((name) => readFileSync(join(root, 'src/toolbar/dialogs', name), 'utf8'));
    const choiceSource = readFileSync(join(root, 'src/toolbar/dialogs/choice.ts'), 'utf8');
    const promptSource = readFileSync(join(root, 'src/toolbar/dialogs/prompt.ts'), 'utf8');

    expect(controlDispatchSource).toContain('okLabel: pageScaleText.ok');
    expect(controlDispatchSource).toContain('invalidMessage: isScale');
    expect(controlDispatchSource).not.toContain("okLabel: 'OK'");
    expect(dynamicDefaultsSource).toContain('thenByLabel: strings.sortThenBy');
    expect(dynamicDefaultsSource).toContain('noThenByLabel: strings.sortNoThenBy');
    expect(dynamicDefaultsSource).toContain('addLevelLabel: strings.sortAddLevel');
    expect(dynamicDefaultsSource).toContain('deleteLevelLabel: strings.sortDeleteLevel');
    expect(dynamicDefaultsSource).toContain('copyLevelLabel: strings.sortCopyLevel');
    expect(dynamicDefaultsSource).toContain('levelUnavailableLabel: strings.sortLevelUnavailable');
    expect(dynamicDefaultsSource).toContain(
      'const invalidRange = strings.advancedFilterInvalidRange',
    );
    expect(dynamicDefaultsSource).toContain('showRenameSheetDialog: (opts) =>');
    expect(dynamicDefaultsSource).toContain('okLabel: strings.hyperlinkDialog.ok');
    expect(dynamicDefaultsSource).toContain('cancelLabel: strings.hyperlinkDialog.cancel');
    expect(dynamicDefaultsSource).toContain(
      'projectFormatToolbar: opts.projectFormatToolbar ?? noop',
    );
    expect(dynamicDefaultsSource).toContain('refreshWorkbookCells:');
    expect(dynamicDefaultsSource).toContain('opts.refreshCells ??');
    expect(toolbarDefaultsSource).toContain('showZoomDialog({');
    expect(toolbarDefaultsSource).toContain('invalidMessage: strings.zoomDialogInvalid');
    expect(toolbarDefaultsSource).not.toContain('showNumberPrompt({');
    for (const source of dialogSources) {
      expect(source).not.toContain("?? 'OK'");
      expect(source).not.toContain("?? 'Cancel'");
    }
    expect(dialogSources.join('\n')).not.toContain("?? 'Then by'");
    expect(dialogSources.join('\n')).not.toContain("?? '(none)'");
    expect(choiceSource).not.toContain("?? 'OK'");
    expect(choiceSource).not.toContain("?? 'Cancel'");
    expect(promptSource).not.toContain("?? 'OK'");
    expect(promptSource).not.toContain("?? 'Cancel'");
    expect(promptSource).not.toContain('Enter a valid number.');
  });

  it('keeps toolbar dialog input focus/select centralized in the shared shell helper', () => {
    const dialogSources = sourceFilesUnder('src/toolbar/dialogs')
      .filter((path) => !path.endsWith('/shell.ts'))
      .map((path) => ({
        path,
        source: readFileSync(path, 'utf8'),
      }));
    const directSelects = dialogSources
      .filter(({ source }) => source.includes('.select();') || source.includes('.select('))
      .map(({ path }) => path.replace(`${root}/`, ''));
    const shellSource = readFileSync(join(root, 'src/toolbar/dialogs/shell.ts'), 'utf8');

    expect(directSelects).toEqual([]);
    expect(shellSource).toContain('export const focusAndSelectInput');
    expect(shellSource).toContain('focusAndSelectInput(input)');
  });

  it('keeps toolbar dialog error row updates centralized in the shared shell helper', () => {
    const dialogSources = sourceFilesUnder('src/toolbar/dialogs')
      .filter((path) => !path.endsWith('/shell.ts'))
      .map((path) => ({
        path,
        source: readFileSync(path, 'utf8'),
      }));
    const directErrorUpdates = dialogSources
      .filter(
        ({ source }) =>
          source.includes('errorRow.textContent') || source.includes('errorRow.hidden = false'),
      )
      .map(({ path }) => path.replace(`${root}/`, ''));
    const shellSource = readFileSync(join(root, 'src/toolbar/dialogs/shell.ts'), 'utf8');

    expect(directErrorUpdates).toEqual([]);
    expect(shellSource).toContain('export const showDialogError');
    expect(shellSource).toContain('export const clearDialogError');
    expect(shellSource).toContain('showDialogError(errorRow, message)');
  });

  it('keeps ribbon menu labels from branching directly on Japanese locale', () => {
    const localeBranchedMenus = sourcesOutsidePrimitives()
      .filter(
        ({ source }) => source.includes("ribbonLang === 'ja' ?") || source.includes('const ja ='),
      )
      .map(({ name }) => name);

    expect(localeBranchedMenus).toEqual([]);
  });

  it('keeps visual tile and swatch DOM centralized in shared visual primitives', () => {
    const directVisualRows = sourcesOutsidePrimitives()
      .filter(
        ({ source }) =>
          /className\s*=\s*['"][^'"]*fc-tb__visual-tile/.test(source) ||
          /className\s*=\s*['"][^'"]*fc-tb__visual-grid/.test(source) ||
          /className\s*=\s*['"][^'"]*fc-tb__color-swatch/.test(source) ||
          /className\s*=\s*['"][^'"]*fc-tb__symbol-tile/.test(source) ||
          /className\s*=\s*['"][^'"]*fc-tb__symbol-grid/.test(source),
      )
      .map(({ name }) => name);

    expect(directVisualRows).toEqual([]);
  });

  it('uses the shared visual tile grid helper for gallery grids', () => {
    const directVisualGrids = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes('visualMenuGrid('))
      .map(({ name }) => name);

    expect(directVisualGrids).toEqual([]);
  });

  it('keeps menu section headings centralized in menuSectionHeader', () => {
    const directHeadings = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes("className = 'fc-tb__menu-heading'"))
      .map(({ name }) => name);

    expect(directHeadings).toEqual([]);
  });

  it('keeps submenu trigger affordances centralized in menuSubmenuTrigger', () => {
    const directSubmenuTriggers = sourcesOutsidePrimitives()
      .filter(
        ({ source }) =>
          source.includes('fc-tb__menu-item--submenu') ||
          source.includes('fc-tb__menu-item__caret') ||
          source.includes("aria-haspopup', 'menu'") ||
          source.includes("aria-expanded', 'false'") ||
          source.includes("setAttribute('aria-controls'"),
      )
      .map(({ name }) => name);

    expect(directSubmenuTriggers).toEqual([]);
  });

  it('keeps submenu trigger panel ownership on the shared controlsId option', () => {
    const sources = new Map(menuSources().map(({ name, source }) => [name, source]));
    const generalSource = sources.get('general.ts');

    expect(generalSource).toContain('opts: { controlsId?: string } = {}');
    expect(generalSource).toContain("button.setAttribute('aria-controls', opts.controlsId)");
    expect(sources.get('conditional.ts')).toContain(
      'menuSubmenuTrigger(btn, { cfSubmenu: key }, { controlsId: cfSubmenuId(key) })',
    );
    expect(sources.get('conditional.ts')).toContain('id: cfSubmenuId(key)');
    expect(sources.get('borders.ts')).toContain(
      'menuSubmenuTrigger(btn, undefined, { controlsId: borderSubmenuId(submenuKey) })',
    );
    expect(sources.get('borders.ts')).toContain("id: borderSubmenuId('lineStyle')");
    expect(sources.get('borders.ts')).toContain("id: borderSubmenuId('lineColor')");
  });

  it('keeps preset icon spacer DOM centralized in menuIconSpacer', () => {
    const directSpacers = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes('fc-tb__menu-item__icon-spacer'))
      .map(({ name }) => name);

    expect(directSpacers).toEqual([]);
  });

  it('keeps submenu item text DOM centralized in submenuItemText', () => {
    const directSubmenuText = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes('fc-tb__submenu-item__text'))
      .map(({ name }) => name);

    expect(directSubmenuText).toEqual([]);
  });

  it('keeps shared primitive span creation centralized in menuSpan', () => {
    const generalSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'general.ts',
    );

    expect(generalSource).toContain('const menuSpan');
    expect(generalSource?.match(/document\.createElement\('span'\)/g) ?? []).toHaveLength(1);
  });

  it('keeps shared primitive div creation centralized in menuDiv', () => {
    const generalSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'general.ts',
    );

    expect(generalSource).toContain('const menuDiv');
    expect(generalSource?.match(/document\.createElement\('div'\)/g) ?? []).toHaveLength(1);
  });

  it('keeps labeled gallery sections centralized in menuLabeledGrid', () => {
    const directLabeledGridDom = sourcesOutsidePrimitives()
      .filter(
        ({ source }) =>
          /className\s*=\s*['"]fc-tb__(?:table|cell)style-heading['"]/.test(source) ||
          /className\s*=\s*['"]fc-tb__(?:table|cell)style-grid['"]/.test(source),
      )
      .map(({ name }) => name);

    expect(directLabeledGridDom).toEqual([]);
    expect(
      new Map(menuSources().map(({ name, source }) => [name, source])).get('styles.ts'),
    ).toContain('menuLabeledGrid(');
  });

  it('keeps gallery scroll bodies centralized in menuScrollBody', () => {
    const stylesSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'styles.ts',
    );

    expect(stylesSource).toContain('menuScrollBody(');
    expect(stylesSource).not.toMatch(/className\s*=\s*['"]fc-tb__tablestyle-scroll['"]/);
  });

  it('keeps table style swatch preview div creation centralized', () => {
    const stylesSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'styles.ts',
    );

    expect(stylesSource).toContain('tableStyleSwatchPart(');
    expect(stylesSource?.match(/document\.createElement\('div'\)/g) ?? []).toHaveLength(1);
  });

  it('keeps cell style chip text DOM centralized in menuTextChip', () => {
    const stylesSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'styles.ts',
    );

    expect(stylesSource).toContain('menuTextChip(');
    expect(stylesSource).not.toContain('.textContent = label');
  });

  it('keeps conditional formatting panel containers centralized in cfPanel', () => {
    const conditionalSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'conditional.ts',
    );

    expect(conditionalSource).toContain('cfPanel(');
    expect(conditionalSource).not.toMatch(
      /className\s*=\s*['"]fc-tb__cf-(?:choice-row|choice-grid-panel|icon-panel)['"]/,
    );
  });

  it('keeps conditional formatting span creation centralized in cfSpan', () => {
    const conditionalSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'conditional.ts',
    );

    expect(conditionalSource).toContain('cfSpan(');
    expect(conditionalSource?.match(/document\.createElement\('span'\)/g) ?? []).toHaveLength(1);
  });

  it('uses the shared preset primitive for specialized preset-row menus', () => {
    const sources = new Map(menuSources().map(({ name, source }) => [name, source]));

    expect(sources.get('borders.ts')).toContain('menuPresetButton(');
    expect(sources.get('conditional.ts')).toContain('menuPresetButton(');
    expect(sources.get('text-orientation.ts')).toContain('menuPresetButton(');
  });

  it('renders Text Orientation menu previews as colored semantic paths', () => {
    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
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

  it('derives dynamic dropdown ids from the shared activation menu map', () => {
    const dynamicDropdownsSource = readFileSync(join(ribbonDir, 'dynamic-dropdowns.ts'), 'utf8');

    expect(dynamicDropdownsSource).toContain('Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND)');
    expect(dynamicDropdownsSource).not.toMatch(
      /const DYNAMIC_RIBBON_DROPDOWN_IDS[\s\S]*new Set\(\[[\s\S]*menu-/,
    );
  });

  it('keeps top-level menu factory ids registered in the shared activation model', () => {
    const registeredMenuIds = new Set([
      ...Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND),
      RIBBON_BORDERS_MENU_ID,
    ]);
    const unregistered = sourcesOutsidePrimitives()
      .flatMap(({ name, source }) =>
        Array.from(source.matchAll(/createMenu\('([^']+)'\)/g)).map(
          (match) => `${name}:${match[1] ?? ''}`,
        ),
      )
      .filter((entry) => !registeredMenuIds.has(entry.split(':')[1] ?? ''))
      .sort();

    expect(unregistered).toEqual([]);
  });

  it('keeps dynamic dropdown refresh routing table-driven', () => {
    const dynamicDropdownsSource = readFileSync(join(ribbonDir, 'dynamic-dropdowns.ts'), 'utf8');

    expect(dynamicDropdownsSource).toContain('DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS');
    expect(dynamicDropdownsSource).toContain('menuRefreshers[spec.menuId]?.(menu)');
    expect(dynamicDropdownsSource).not.toContain("if (spec.menuId === 'menu-");
    const registeredMenuIds = new Set(Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND));
    const refresherMenuIds = Object.keys(DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS);

    expect(refresherMenuIds.filter((id) => !registeredMenuIds.has(id))).toEqual([]);
  });

  it('keeps every dynamic dropdown update hook routed through menuRefreshers', () => {
    const noopCtx: Pick<DynamicDropdownsCtx, keyof DynamicDropdownsCtx> = {
      applyRibbonPasteAction: vi.fn(),
      applyPivotTableAction: vi.fn(),
      applyDefinedNameAction: vi.fn(),
      applyLinksAction: vi.fn(),
      applyCopyAction: vi.fn(),
      applyFillSeries: vi.fn(),
      applyFillDirection: vi.fn(),
      applyClearAction: vi.fn(),
      applyUnderlineAction: vi.fn(),
      applyWrapAction: vi.fn(),
      applyMergeAction: vi.fn(),
      applyFreezeAction: vi.fn(),
      applyTextOrientationAction: vi.fn(),
      applyCellInsertAction: vi.fn(),
      applyCellDeleteAction: vi.fn(),
      applyCellFormatAction: vi.fn(),
      applyPageBreakAction: vi.fn(),
      applySheetBackgroundAction: vi.fn(),
      applyPrintAreaAction: vi.fn(),
      applyArrangeAction: vi.fn(),
      applyUiTheme: vi.fn(),
      focusSheet: vi.fn(),
      applySortMenuAction: vi.fn(),
      applyFindSelectAction: vi.fn(),
      applyAutoSumFormula: vi.fn(),
      applyFormulaAuditAction: vi.fn(),
      applyWatchAction: vi.fn(),
      applyReviewCommentAction: vi.fn(),
      applyProtectAction: vi.fn(),
      applyCalcOptionAction: vi.fn(),
      createRecommendedChartFromSelection: vi.fn(),
      createChartFromSelection: vi.fn(),
      chartKindFromAction: vi.fn(),
      insertPictureFromRibbon: vi.fn(),
      insertShapeFromRibbon: vi.fn(),
      insertScreenshotFromRibbon: vi.fn(),
      applyScriptAction: vi.fn(),
      applyPdfAction: vi.fn(),
      createTableFromSelection: vi.fn(),
      openTableStyleFooterAction: vi.fn(),
      applyPivotTableStyleFromRibbon: vi.fn(),
      applyCellStyleFromRibbon: vi.fn(),
      openCellStyleFooterAction: vi.fn(),
      applyCurrencyPreset: vi.fn(),
      openCurrencyFooterAction: vi.fn(),
      splitTextToColumns: vi.fn(),
      splitTextToColumnsCustom: vi.fn(),
      applyDataValidationAction: vi.fn(),
      applyAddInAction: vi.fn(),
      applyConditionalMenuAction: vi.fn(),
      applySymbolAction: vi.fn(),
      getInst: vi.fn(),
      updateCalcOptionsMenu: vi.fn(),
      updateCellDeleteMenu: vi.fn(),
      updateCellInsertMenu: vi.fn(),
      updateCellStylesMenu: vi.fn(),
      updateClearMenu: vi.fn(),
      updateClearArrowsMenu: vi.fn(),
      updateCurrencyMenu: vi.fn(),
      updateDataValidationMenu: vi.fn(),
      updateDefinedNamesMenu: vi.fn(),
      updateErrorCheckingMenu: vi.fn(),
      updateFillMenu: vi.fn(),
      updateFormatCellsMenu: vi.fn(),
      updateFreezeMenu: vi.fn(),
      updateLinksMenu: vi.fn(),
      updatePasteMenu: vi.fn(),
      updateArrangeMenu: vi.fn(),
      updatePageBreaksMenu: vi.fn(),
      updatePrintAreaMenu: vi.fn(),
      updateProtectMenu: vi.fn(),
      updatePageThemeMenu: vi.fn(),
      updateReviewCommentsMenu: vi.fn(),
      updateSortMenu: vi.fn(),
      updateTableStylesMenu: vi.fn(),
      updateTextOrientationMenu: vi.fn(),
      updateWatchMenu: vi.fn(),
      closeBorderMenu: vi.fn(),
      closeFreezeMenu: vi.fn(),
      closePrintAreaMenu: vi.fn(),
      closeSymbolMenu: vi.fn(),
      getConditionalMenu: vi.fn(),
    };
    const updateHooks = Object.keys(noopCtx)
      .filter((key) => /^update[A-Za-z0-9]+Menu$/.test(key))
      .sort();
    const routedHooks = Array.from(
      new Set(Object.values(DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS)),
    ).sort();

    expect(routedHooks).toEqual(updateHooks);
  });

  it('keeps dynamic dropdown handler dataset keys derived from the shared handler attrs', () => {
    const datasetKeyForAttr = (attr: string): string =>
      attr.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
    const expected = new Set([
      ...DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS.map(datasetKeyForAttr),
      'cfAction',
      'cfSubmenu',
      'formatSubmenu',
    ]);

    expect(DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS).toEqual(expected);
  });

  it('keeps dynamic dropdown manifests exported from the public entrypoint', () => {
    const indexSource = readFileSync(join(root, 'src/index.ts'), 'utf8');

    for (const symbol of [
      'DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS',
      'DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS',
      'DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS',
      'DynamicDropdownMenuRefresherKey',
    ]) {
      expect(indexSource, symbol).toContain(symbol);
    }
  });

  it('keeps public shared dialog helpers documented for host and wrapper reuse', () => {
    const readmePaths = [
      join(root, '../..', 'README.md'),
      join(root, '../..', 'README_ja.md'),
      join(root, 'README.md'),
      join(root, '../formulon-cell-react/README.md'),
      join(root, '../formulon-cell-react/README_ja.md'),
      join(root, '../formulon-cell-vue/README.md'),
      join(root, '../formulon-cell-vue/README_ja.md'),
    ];
    for (const readmePath of readmePaths) {
      const source = readFileSync(readmePath, 'utf8');
      expect(source, readmePath).toContain('reportDialogLabels');
      expect(source, readmePath).toContain('showReport');
      expect(source, readmePath).toContain('projectDisabledReason');
    }
  });

  it('keeps dynamic dropdown dispatcher attrs aligned with registered handlers', () => {
    const dynamicDropdownsSource = readFileSync(join(ribbonDir, 'dynamic-dropdowns.ts'), 'utf8');
    const handlersBlock = dynamicDropdownsSource.match(
      /const DYNAMIC_DROPDOWN_HANDLERS:[\s\S]*?= \[([\s\S]*?)\n {2}\];/,
    )?.[1];
    const handlerAttrs = Array.from(handlersBlock?.matchAll(/attr: '([^']+)'/g) ?? []).flatMap(
      (match) => (match[1] ? [match[1]] : []),
    );

    expect(handlerAttrs).toEqual(DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS);
  });

  it('keeps dynamic dropdown event target and disabled checks centralized', () => {
    const dynamicDropdownsSource = readFileSync(join(ribbonDir, 'dynamic-dropdowns.ts'), 'utf8');

    expect(dynamicDropdownsSource).toContain('const eventElement');
    expect(dynamicDropdownsSource).toContain('const isDisabledMenuControl');
    expect(
      dynamicDropdownsSource.split('\n').filter((line) => line.includes('event.target')),
    ).toHaveLength(1);
    expect(
      dynamicDropdownsSource.split('\n').filter((line) => line.includes('.disabled')),
    ).toHaveLength(1);
  });

  it('keeps dynamic dropdown viewport clamp and scroll projection centralized', () => {
    const dynamicDropdownsSource = readFileSync(join(ribbonDir, 'dynamic-dropdowns.ts'), 'utf8');

    expect(dynamicDropdownsSource).toContain('const applyVerticalViewportLimit');
    expect(dynamicDropdownsSource).toContain('viewportSize()');
    expect(dynamicDropdownsSource).toContain('import { clamp, viewportSize }');
    expect(dynamicDropdownsSource.match(/window\.innerWidth/g) ?? []).toHaveLength(0);
    expect(dynamicDropdownsSource.match(/window\.innerHeight/g) ?? []).toHaveLength(0);
    expect(dynamicDropdownsSource.match(/style\.overflowY/g) ?? []).toHaveLength(2);
    expect(dynamicDropdownsSource.match(/style\.overscrollBehavior/g) ?? []).toHaveLength(2);
    expect(dynamicDropdownsSource.match(/applyVerticalViewportLimit\(/g) ?? []).toHaveLength(2);
  });

  it('keeps body-attached overlay viewport sizing centralized in overlay-position', () => {
    const allowedFiles = new Set(['src/interact/overlay-position.ts']);
    const files = ['src/interact', 'src/mount', 'src/toolbar', 'src/components'].flatMap(
      sourceFilesUnder,
    );
    const violations: string[] = [];

    for (const file of files) {
      if (allowedFiles.has(file)) continue;
      const source = readFileSync(join(root, file), 'utf8');
      const lines = source.split('\n');
      lines.forEach((line, index) => {
        if (/window\.inner(?:Width|Height)/.test(line)) {
          violations.push(`${file}:${index + 1}: ${line.trim()}`);
        }
      });
    }

    expect(violations).toEqual([]);
  });

  it('keeps menu disabled state projection centralized in dropdown defaults', () => {
    const defaultsSource = readFileSync(join(mountDir, 'dynamic-dropdowns-defaults.ts'), 'utf8');

    expect(defaultsSource).toContain('const setMenuControlDisabled');
    expect(defaultsSource).toContain('projectDisabledState(button, disabled');
    expect(defaultsSource.match(/\.disabled\s*=/g) ?? []).toHaveLength(0);
    expect(defaultsSource.match(/setAttribute\('aria-disabled'/g) ?? []).toHaveLength(0);
    expect(defaultsSource).not.toContain("setAttribute('aria-description'");
    expect(defaultsSource).not.toContain('dataset.menuDisabledReason');
  });

  it('keeps disabled control state mutation centralized in menu-a11y', () => {
    const allowedFiles = new Set(['src/toolbar/menu-a11y.ts']);
    const files = disabledStateAuditDirs.flatMap(sourceFilesUnder);
    const violations: string[] = [];

    for (const file of files) {
      if (allowedFiles.has(file)) continue;
      const source = readFileSync(join(root, file), 'utf8');
      const lines = source.split('\n');
      lines.forEach((line, index) => {
        if (/\.disabled\s*=(?!=)/.test(line) || /setAttribute\((['"])aria-disabled\1/.test(line)) {
          violations.push(`${file}:${index + 1}: ${line.trim()}`);
        }
      });
    }

    expect(violations).toEqual([]);
  });

  it('keeps shared menu a11y disabled checks aligned with aria-disabled', () => {
    const menuA11ySource = readFileSync(join(root, 'src/toolbar/menu-a11y.ts'), 'utf8');

    expect(menuA11ySource).toContain('!item.disabled');
    expect(menuA11ySource).toContain("item.getAttribute('aria-disabled') !== 'true'");
  });

  it('projects disabled reasons through the shared helper', () => {
    const button = document.createElement('button');
    projectDisabledReason(button, 'Unavailable', { datasetKey: 'menuDisabledReason' });
    expect(button.title).toBe('Unavailable');
    expect(button.getAttribute('aria-description')).toBe('Unavailable');
    expect(button.dataset.menuDisabledReason).toBe('Unavailable');

    projectDisabledReason(button, null, { datasetKey: 'menuDisabledReason' });
    expect(button.title).toBe('');
    expect(button.getAttribute('aria-description')).toBeNull();
    expect(button.dataset.menuDisabledReason).toBeUndefined();

    projectDisabledReason(button, 'Coming soon', {
      datasetKey: 'ribbonDisabledReason',
      titlePrefix: 'Automate',
    });
    expect(button.title).toBe('Automate\nComing soon');
    expect(button.getAttribute('aria-description')).toBe('Coming soon');
    expect(button.dataset.ribbonDisabledReason).toBe('Coming soon');
    projectDisabledReason(button, null, {
      datasetKey: 'ribbonDisabledReason',
      titlePrefix: 'Automate',
    });
    expect(button.title).toBe('Automate');
    expect(button.dataset.ribbonDisabledReason).toBeUndefined();

    const input = document.createElement('input');
    projectDisabledReason(input, 'Read only', {
      ariaDescription: false,
      describedById: 'readonly-note',
    });
    expect(input.title).toBe('Read only');
    expect(input.getAttribute('aria-describedby')).toBe('readonly-note');
    expect(input.getAttribute('aria-description')).toBeNull();
    projectDisabledReason(input, null, {
      ariaDescription: false,
      describedById: 'readonly-note',
    });
    expect(input.title).toBe('');
    expect(input.getAttribute('aria-describedby')).toBeNull();
  });

  it('projects disabled control state through the shared helper', () => {
    const button = document.createElement('button');
    projectDisabledState(button, true, 'Unavailable', {
      datasetKey: 'disabledReason',
      titlePrefix: 'Insert Function',
    });

    expect(button.disabled).toBe(true);
    expect(button.getAttribute('aria-disabled')).toBe('true');
    expect(button.title).toBe('Insert Function\nUnavailable');
    expect(button.dataset.disabledReason).toBe('Unavailable');

    projectDisabledState(button, false, 'Unavailable', {
      datasetKey: 'disabledReason',
      titlePrefix: 'Insert Function',
    });

    expect(button.disabled).toBe(false);
    expect(button.getAttribute('aria-disabled')).toBe('false');
    expect(button.title).toBe('Insert Function');
    expect(button.dataset.disabledReason).toBeUndefined();
  });

  it('keeps range picker collapsed dialog styling centralized', () => {
    const frameSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/frame.css'),
      'utf8',
    );
    const controlsSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/controls.css'),
      'utf8',
    );

    expect(frameSource).toContain('.fc-fmtdlg--range-picking');
    expect(frameSource).toContain('pointer-events: none');
    expect(frameSource).toContain('width: min(460px, calc(100vw - 24px))');
    expect(frameSource).toContain(':not(:has(.fc-range-picker--picking))');
    expect(controlsSource).toContain('.fc-range-picker--picking > input');
    expect(controlsSource).toContain('.fc-range-picker--picking .fc-range-picker__btn');
  });

  it('keeps Format Cells chrome close to Japanese Excel 365 desktop', () => {
    const themeSource = readFileSync(join(root, 'src/styles/theme-paper.css'), 'utf8');
    const frameSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/frame.css'),
      'utf8',
    );
    const controlsSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/controls.css'),
      'utf8',
    );
    const choicesSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/choices.css'),
      'utf8',
    );
    const tabsContentSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/tabs-content.css'),
      'utf8',
    );
    const numberSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/number.css'),
      'utf8',
    );
    const swatchesSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/swatches-and-lines.css'),
      'utf8',
    );
    const bordersSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/borders.css'),
      'utf8',
    );
    const customSelectSource = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/custom-select.css'),
      'utf8',
    );

    expect(themeSource).toContain('--fc-fmtdlg-tab-active-bg: #107c41');
    expect(themeSource).toContain('--fc-fmtdlg-tab-active-color: #ffffff');
    expect(themeSource).toContain('--fc-fmtdlg-tab-active-underline: none');
    expect(themeSource).toContain('--fc-fmtdlg-list-hover-bg: #eeeeee');
    expect(frameSource).toContain('border-radius: var(--fc-fmtdlg-tab-radius, 0)');
    expect(frameSource).toContain('filter: none');
    expect(controlsSource).toContain('border-radius: 4px');
    expect(controlsSource).toContain('box-shadow: 0 0 0 1px var(--fc-fmtdlg-list-focus-border)');
    expect(controlsSource).not.toContain('box-shadow: 0 0 0 2px var(--fc-accent-soft)');
    expect(controlsSource).toMatch(
      /\.fc-range-picker__btn::after\s*\{[\s\S]*?width: 8px;[\s\S]*?height: 8px;[\s\S]*?linear-gradient\(45deg,[\s\S]*?#185abd[\s\S]*?content: "";/,
    );
    expect(controlsSource).not.toContain('content: "↗"');
    expect(choicesSource).toContain('border-radius: 4px');
    expect(tabsContentSource).toContain(
      'box-shadow: inset 0 0 0 1px var(--fc-fmtdlg-list-focus-border)',
    );
    expect(tabsContentSource).toContain(
      '.fc-fmtdlg__font-list-item:hover {\n    background: var(--fc-fmtdlg-list-hover-bg);',
    );
    expect(numberSource).toContain(
      'box-shadow: inset 0 0 0 1px var(--fc-fmtdlg-list-focus-border)',
    );
    expect(swatchesSource).toContain('border-radius: 2px');
    expect(swatchesSource).toContain('border: 1px solid var(--fc-fmtdlg-input-hover-border)');
    expect(swatchesSource).toContain('outline: 1px solid var(--fc-accent, currentColor)');
    expect(bordersSource).toContain('border-radius: 2px');
    expect(customSelectSource).toMatch(/\.fc-select__button\s*\{[\s\S]*?border-radius: 2px;/);
    expect(customSelectSource).toContain('box-shadow: none');
    expect(customSelectSource).toMatch(/\.fc-select__list\s*\{[\s\S]*?border-radius: 2px;/);
    expect(customSelectSource).toMatch(/\.fc-select__option\s*\{[\s\S]*?border-radius: 0;/);
    expect(customSelectSource).not.toContain('box-shadow: 0 0 0 2px var(--fc-accent-soft');
  });

  it('keeps Custom Sort level grid styling in a shared dialog module', () => {
    const appSource = readFileSync(join(root, 'src/styles/core/app.css'), 'utf8');
    const sortSource = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/sort.css'),
      'utf8',
    );

    expect(appSource).toContain('@import "./app/dialog-modules/sort.css"');
    expect(sortSource).toContain('.fc-sortdlg__toolbar');
    expect(sortSource).toContain('.fc-sortdlg__grid-head');
    expect(sortSource).toContain('.fc-sortdlg__levels');
    expect(sortSource).toContain('.fc-sortdlg__level--selected');
    expect(sortSource).toContain('grid-template-columns');
    expect(sortSource).toContain('@media (max-width: 560px)');
  });

  it('keeps Remove Duplicates column checklist styling in a shared dialog module', () => {
    const appSource = readFileSync(join(root, 'src/styles/core/app.css'), 'utf8');
    const source = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/remove-duplicates.css'),
      'utf8',
    );

    expect(appSource).toContain('@import "./app/dialog-modules/remove-duplicates.css"');
    expect(source).toContain('.fc-dedupedlg__actions');
    expect(source).toContain('.fc-dedupedlg__column-list');
    expect(source).toContain('.fc-dedupedlg__column');
    expect(source).toContain('grid-template-columns');
    expect(source).toContain('@media (max-width: 520px)');
  });

  it('keeps toolbar dialog action buttons centralized in the shared shell', () => {
    for (const file of [
      'src/toolbar/dialogs/prompt.ts',
      'src/toolbar/dialogs/report.ts',
      'src/toolbar/dialogs/remove-duplicates.ts',
      'src/toolbar/dialogs/sort.ts',
    ]) {
      const source = readFileSync(join(root, file), 'utf8');
      expect(source).toContain('appendDialogButton(');
      expect(source).not.toContain("const okBtn = document.createElement('button')");
      expect(source).not.toContain("const closeBtn = document.createElement('button')");
      expect(source).not.toContain("const selectAllBtn = document.createElement('button')");
      expect(source).not.toContain("const unselectAllBtn = document.createElement('button')");
      expect(source).not.toContain("const addLevelBtn = document.createElement('button')");
      expect(source).not.toContain("const deleteLevelBtn = document.createElement('button')");
      expect(source).not.toContain("const copyLevelBtn = document.createElement('button')");
    }
  });

  it('keeps toolbar dialog choice buttons centralized in the shared shell', () => {
    const shellSource = readFileSync(join(root, 'src/toolbar/dialogs/shell.ts'), 'utf8');
    const symbolSource = readFileSync(join(root, 'src/toolbar/dialogs/symbol.ts'), 'utf8');

    expect(shellSource).toContain('createDialogChoiceButton');
    expect(shellSource).toContain("button.className = opts.className ?? 'fc-tb__cf-choice'");
    expect(symbolSource).toContain('createDialogChoiceButton({ label: symbol');
    expect(symbolSource).not.toContain("const button = document.createElement('button')");
    expect(symbolSource).not.toContain("button.className = 'fc-tb__cf-choice'");
  });

  it('keeps Text to Columns wizard styling in a shared dialog module', () => {
    const appSource = readFileSync(join(root, 'src/styles/core/app.css'), 'utf8');
    const source = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/text-to-columns.css'),
      'utf8',
    );

    expect(appSource).toContain('@import "./app/dialog-modules/text-to-columns.css"');
    expect(source).toContain('.fc-textcols__types');
    expect(source).toContain('.fc-textcols__delimiter-grid');
    expect(source).toContain('.fc-textcols__preview');
    expect(source).toContain('grid-template-columns');
    expect(source).toContain('@media (max-width: 520px)');
  });

  it('keeps Advanced Filter range form styling in a shared dialog module', () => {
    const appSource = readFileSync(join(root, 'src/styles/core/app.css'), 'utf8');
    const source = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/advanced-filter.css'),
      'utf8',
    );

    expect(appSource).toContain('@import "./app/dialog-modules/advanced-filter.css"');
    expect(source).toContain('.fc-advfilter__ranges');
    expect(source).toContain('.fc-advfilter__row');
    expect(source).toContain('.fc-advfilter__option');
    expect(source).toContain('grid-template-columns');
    expect(source).toContain('@media (max-width: 520px)');
  });

  it('skips aria-disabled menu buttons during shared roving focus', () => {
    const menu = createMenu('menu-a11y-disabled-test');
    const ariaDisabled = menuIconButton('Disabled', 'clear', 'formats', 'clear-formats');
    const enabled = menuIconButton('Enabled', 'clear', 'contents', 'clear-contents');
    ariaDisabled.setAttribute('aria-disabled', 'true');
    menu.append(ariaDisabled, enabled);
    document.body.appendChild(menu);
    prepareMenu(menu);

    focusMenuItem(menu);
    expect(document.activeElement).toBe(enabled);
    expect(ariaDisabled.tabIndex).toBe(-1);
    expect(enabled.tabIndex).toBe(0);

    const key = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    const clickDisabled = vi.fn();
    const clickEnabled = vi.fn();
    ariaDisabled.addEventListener('click', clickDisabled);
    enabled.addEventListener('click', clickEnabled);
    Object.defineProperty(key, 'target', { value: menu });
    handleMenuKeydown(key, menu, { close: vi.fn() });
    expect(clickDisabled).not.toHaveBeenCalled();
    expect(clickEnabled).toHaveBeenCalledTimes(1);
    menu.remove();
  });

  it('applies the common button contract across shared menu primitives', () => {
    const leading = document.createElement('span');
    leading.className = 'test-leading';
    const buttons = [
      { button: menuIconButton('Clear', 'clear', 'formats', 'clear-formats'), key: 'clear' },
      {
        button: menuPresetButton('Bottom', 'borderPreset', 'bottom', leading),
        key: 'borderPreset',
      },
      {
        button: menuPresetButton('No icon', 'borderPreset', 'none', menuIconSpacer()),
        key: 'borderPreset',
      },
      {
        button: colorSwatchButton({
          label: 'Yellow',
          attr: 'fillColor',
          value: '#ffff00',
          color: '#ffff00',
        }),
        key: 'fillColor',
        label: 'Yellow',
      },
      {
        button: menuTextChip({
          label: 'Good',
          attr: 'cellStyle',
          value: 'good',
          className: 'fc-tb__menu-item fc-tb__cellstyle-chip',
        }),
        key: 'cellStyle',
        label: 'Good',
      },
      { button: symbolMenuTile('π'), key: 'symbol', label: 'π' },
      {
        button: visualMenuTile({
          label: 'Column',
          attr: 'chartInsert',
          value: 'column',
          icon: 'chart-column',
        }),
        key: 'chartInsert',
        label: 'Column',
      },
    ];

    for (const { button, key, label } of buttons) {
      expect(button.type).toBe('button');
      expect(button.getAttribute('role')).toBe('menuitem');
      expect(button.dataset[key]).toBeTruthy();
      if (label) {
        expect(button.title).toBe(label);
        expect(button.getAttribute('aria-label')).toBe(label);
      }
    }
  });

  it('creates the base menu button contract used by specialized primitives', () => {
    const button = createMenuButton({
      className: 'fc-tb__menu-item fc-tb__menu-item--custom',
      attr: 'sampleAction',
      value: 'run',
      title: 'Run sample',
      ariaLabel: 'Run sample',
    });

    expect(button.className).toBe('fc-tb__menu-item fc-tb__menu-item--custom');
    expect(button.type).toBe('button');
    expect(button.getAttribute('role')).toBe('menuitem');
    expect(button.dataset.sampleAction).toBe('run');
    expect(button.title).toBe('Run sample');
    expect(button.getAttribute('aria-label')).toBe('Run sample');
  });

  it('creates menu div primitives with shared class and accessibility contracts', () => {
    const menu = createMenu('menu-test');
    expect(menu.id).toBe('menu-test');
    expect(menu.className).toBe('fc-tb__menu');
    expect(menu.hidden).toBe(true);

    const colorGrid = colorSwatchGrid('test-colors');
    expect(colorGrid.className).toBe('fc-tb__color-swatch-grid test-colors');
    expect(colorGrid.getAttribute('role')).toBe('presentation');

    const symbolGrid = symbolMenuGrid('Greek', ['π']);
    expect(symbolGrid.className).toBe('fc-tb__symbol-grid');
    expect(symbolGrid.getAttribute('role')).toBe('presentation');
    expect(symbolGrid.getAttribute('aria-label')).toBe('Greek');
    expect(symbolGrid.querySelectorAll('button')).toHaveLength(1);

    const visualGrid = visualMenuGrid('test-visuals');
    expect(visualGrid.className).toBe('fc-tb__visual-grid test-visuals');
    expect(visualGrid.getAttribute('role')).toBe('presentation');

    const separator = menuSeparator();
    expect(separator.className).toBe('fc-tb__menu-sep');
    expect(separator.getAttribute('role')).toBe('separator');

    const heading = menuSectionHeader('Styles');
    expect(heading.className).toBe('fc-tb__menu-heading');
    expect(heading.getAttribute('role')).toBe('presentation');
    expect(heading.textContent).toBe('Styles');

    const [labeledHeading, labeledGrid] = menuLabeledGrid({
      label: 'Light',
      headingClassName: 'fc-tb__tablestyle-heading',
      gridClassName: 'fc-tb__tablestyle-grid',
      children: [],
    });
    expect(labeledHeading.className).toBe('fc-tb__tablestyle-heading');
    expect(labeledHeading.textContent).toBe('Light');
    expect(labeledGrid.className).toBe('fc-tb__tablestyle-grid');
    expect(labeledGrid.getAttribute('role')).toBe('group');
    expect(labeledGrid.getAttribute('aria-label')).toBe('Light');
  });

  it('creates nested submenus with the shared menu contract', () => {
    const submenu = createSubmenu({
      id: 'menu-test-submenu',
      className: 'fc-tb__submenu fc-tb__submenu--test',
      label: 'Test submenu',
      dataset: { cfPanel: 'highlight' },
    });

    expect(submenu.id).toBe('menu-test-submenu');
    expect(submenu.className).toBe('fc-tb__submenu fc-tb__submenu--test');
    expect(submenu.getAttribute('role')).toBe('menu');
    expect(submenu.getAttribute('aria-label')).toBe('Test submenu');
    expect(submenu.hidden).toBe(true);
    expect(submenu.dataset.cfPanel).toBe('highlight');
  });

  it('decorates submenu triggers with caret and shared accessibility attributes', () => {
    const button = menuPresetButton(
      'Highlight Cells Rules',
      'cfAction',
      'submenu-highlight',
      document.createElement('span'),
    );
    const trigger = menuSubmenuTrigger(
      button,
      { cfSubmenu: 'highlight' },
      { controlsId: 'menu-conditional-highlight' },
    );

    expect(trigger).toBe(button);
    expect(trigger.classList.contains('fc-tb__menu-item--submenu')).toBe(true);
    expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    expect(trigger.getAttribute('aria-controls')).toBe('menu-conditional-highlight');
    expect(trigger.dataset.cfSubmenu).toBe('highlight');
    const caret = trigger.querySelector<HTMLElement>('.fc-tb__menu-item__caret');
    expect(caret?.textContent).toBe('');
    expect(caret?.getAttribute('aria-hidden')).toBe('true');

    const menusCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/menus.css'), 'utf8');
    expect(menusCss).toMatch(
      /\.fc-tb__menu-item__caret\s*\{[\s\S]*?border-top: 4px solid transparent;[\s\S]*?border-bottom: 4px solid transparent;[\s\S]*?border-left: 5px solid var\(--fc-tb-fg-soft\);/,
    );
  });

  it('creates shared preset icon spacers', () => {
    const spacer = menuIconSpacer();

    expect(spacer.tagName).toBe('SPAN');
    expect(spacer.className).toBe('fc-tb__menu-item__icon-spacer');
  });

  it('creates shared submenu item text labels', () => {
    const text = submenuItemText('None');

    expect(text.tagName).toBe('SPAN');
    expect(text.className).toBe('fc-tb__submenu-item__text');
    expect(text.textContent).toBe('None');
  });
});
