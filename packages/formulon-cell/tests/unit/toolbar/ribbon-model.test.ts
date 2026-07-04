import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  backstageMenuText,
  buildRibbonModel,
  conditionalMenuText,
  EXCEL365_STANDARD_RIBBON_TABS,
  excelRibbonIconPaths,
  fluentIconPaths,
  HOME_MIXED_LAYOUT_GROUP_VARIANTS,
  HOME_STACKED_LAYOUT_GROUP_VARIANTS,
  HOME_TILE_LAYOUT_GROUP_VARIANTS,
  OPTIONAL_RIBBON_TABS,
  pageScaleMenuText,
  ribbonActivatableCommandIds,
  ribbonActivatableSurfaceCommandIds,
  ribbonCommandIds,
  ribbonCommands,
  ribbonDisplayText,
  ribbonSurfaceCommandIds,
  ribbonTabCommandIds,
  toolbarMenuText,
  viewToggleMenuText,
} from '../../../src/index.js';

describe('toolbar/ribbon-model', () => {
  const ribbonStylesDir = join(process.cwd(), 'src/styles/toolbar/ribbon');
  const toolbarBaseDir = join(process.cwd(), 'src/styles/toolbar/base');

  it('keeps the shared core ribbon command surface unique within each tab', () => {
    const keys = buildRibbonModel('en').flatMap((tab) =>
      tab.groups.flatMap((group) =>
        group.commands.map((command) => `${tab.id}:${command.id}:${command.kind}`),
      ),
    );
    const duplicates = keys.filter((key, index) => keys.indexOf(key) !== index);

    expect(duplicates).toEqual([]);
  });

  it('keeps command surfaces locale-independent', () => {
    expect(ribbonCommandIds('ja')).toEqual(ribbonCommandIds('en'));
    expect(ribbonActivatableCommandIds('ja')).toEqual(ribbonActivatableCommandIds('en'));
    expect(ribbonSurfaceCommandIds()).toEqual(ribbonCommandIds('en'));
    expect(ribbonActivatableSurfaceCommandIds()).toEqual(ribbonActivatableCommandIds('en'));
    expect(ribbonSurfaceCommandIds({ tabs: EXCEL365_STANDARD_RIBBON_TABS })).toEqual(
      ribbonCommandIds('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS }),
    );
    expect(ribbonActivatableSurfaceCommandIds({ tabs: EXCEL365_STANDARD_RIBBON_TABS })).toEqual(
      ribbonActivatableCommandIds('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS }),
    );
  });

  it('keeps the Excel 365 Home tab group and command placement explicit', () => {
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');

    expect(home?.groups.map((group) => group.variant)).toEqual([
      'clipboard',
      'font',
      'alignment',
      'number',
      'styles',
      'cells',
      'editing',
    ]);

    const commandsByGroup = new Map(
      home?.groups.map((group) => [group.variant, group.commands.map((command) => command.id)]) ??
        [],
    );

    expect(commandsByGroup.get('clipboard')).toEqual(['paste', 'cut', 'copy', 'formatPainter']);
    expect(commandsByGroup.get('font')).toEqual([
      'fontFamily',
      'fontSize',
      'fontGrow',
      'fontShrink',
      'font-row-2',
      'bold',
      'italic',
      'underline',
      'strike',
      'borders',
      'fillColor',
      'fontColor',
    ]);
    expect(commandsByGroup.get('alignment')).toEqual([
      'top',
      'middle',
      'bottomAlign',
      'alignL',
      'alignC',
      'alignR',
      'alignment-row-2',
      'textOrientation',
      'wrap',
      'indentDecrease',
      'indentIncrease',
      'merge',
    ]);
    expect(commandsByGroup.get('number')).toEqual([
      'numberFormat',
      'number-row-2',
      'currency',
      'percent',
      'comma',
      'decDown',
      'decUp',
    ]);
    expect(commandsByGroup.get('styles')).toEqual(['conditional', 'formatTableHome', 'cellStyles']);
    expect(commandsByGroup.get('cells')).toEqual(['insertRows', 'deleteRows', 'formatCellsHome']);
    expect(commandsByGroup.get('editing')).toEqual([
      'autosum',
      'fillHome',
      'clearFormat',
      'sortFilterHome',
      'findHome',
    ]);
  });

  it('keeps Home wide-button groups on the shared tile layout CSS', () => {
    const groupsCss = readFileSync(join(ribbonStylesDir, 'groups.css'), 'utf8');

    for (const variant of HOME_TILE_LAYOUT_GROUP_VARIANTS) {
      expect(groupsCss).toContain(`.demo__ribbon-group--${variant} .demo__ribbon-tools`);
    }
    expect(groupsCss).toContain('.demo__ribbon-group--tiles .demo__rb--wide');
    expect(groupsCss).toContain('.demo__ribbon-group--tiles .demo__rb-icon');
    expect(groupsCss).toContain('flex-wrap: nowrap;');
    expect(groupsCss).toContain('height: 66px;');
    expect(groupsCss).toContain('grid-template-columns: 58px repeat(2, 30px);');
    expect(groupsCss).toContain('.demo__ribbon-group--tiles .demo__rb .demo__rb-split-chevron');
    expect(groupsCss).toContain('position: absolute;');
    expect(groupsCss).toContain('bottom: 5px;');
    expect(groupsCss).toMatch(
      /\.demo__ribbon-group--stacked \.demo__rb--stacked \.demo__rb-icon,[\s\S]*?width: 17px;[\s\S]*?height: 17px;/,
    );
  });

  it('loads group-specific ribbon CSS after generic button sizing', () => {
    for (const cssFile of ['../ribbon.css', '../../toolbar.css']) {
      const css = readFileSync(join(ribbonStylesDir, cssFile), 'utf8');
      const generic = css.indexOf('layout-and-buttons.css');
      const groups = css.indexOf('groups.css');
      expect(generic, cssFile).toBeGreaterThanOrEqual(0);
      expect(groups, cssFile).toBeGreaterThan(generic);
    }
  });

  it('keeps ribbon chrome close to the Excel 365 white desktop surface', () => {
    const tokensCss = readFileSync(join(ribbonStylesDir, '../base/tokens.css'), 'utf8');
    const layoutCss = readFileSync(join(ribbonStylesDir, 'layout-and-buttons.css'), 'utf8');

    expect(tokensCss).toContain('--fc-tb-ribbon-bg: #ffffff;');
    expect(tokensCss).toContain('--fc-tb-ribbon-rail: #faf9f8;');
    expect(tokensCss).toContain('--fc-tb-ribbon-hover: #f3f2f1;');
    expect(tokensCss).toContain('--fc-tb-ribbon-pressed: #edebe9;');
    expect(tokensCss).toContain('--fc-tb-ribbon-line: #e1dfdd;');
    expect(layoutCss).toContain('.demo__ribbon-tabs');
    expect(layoutCss).toContain('background: var(--fc-tb-ribbon-bg);');
    expect(layoutCss).toMatch(
      /\.demo__ribbon:not\(\[hidden\]\)\s*\{[\s\S]*?min-height: 96px;[\s\S]*?padding: 3px 6px 2px;/,
    );
    expect(layoutCss).toMatch(
      /\.demo__ribbon-tab\s*\{[\s\S]*?font-weight: 400;[\s\S]*?line-height: 31px;/,
    );
    expect(layoutCss).toMatch(/\.demo__ribbon-tab--active\s*\{[\s\S]*?font-weight: 600;/);
    expect(layoutCss).toMatch(
      /\.demo__ribbon-label\s*\{[\s\S]*?height: 14px;[\s\S]*?color: color-mix\(in srgb, var\(--fc-tb-fg-soft\) 82%, var\(--fc-tb-ribbon-bg\)\);[\s\S]*?line-height: 14px;/,
    );
    expect(layoutCss).toMatch(
      /\.demo__ribbon-group\s*\{[\s\S]*?padding: 0 6px;[\s\S]*?border-right: 1px solid var\(--fc-tb-ribbon-line\);/,
    );
    expect(layoutCss).toContain('height: 3px;');
    expect(layoutCss).toContain(
      'background: var(--fc-tb-ribbon-pressed, var(--fc-tb-ribbon-hover));',
    );
    expect(layoutCss).toMatch(
      /\.demo__rb--active\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-pressed, var\(--fc-tb-ribbon-hover\)\);[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);/,
    );
    expect(layoutCss).toMatch(
      /\.demo__rb\s*\{[\s\S]*?width: 30px;[\s\S]*?height: 26px;[\s\S]*?border-radius: 2px;/,
    );
    expect(layoutCss).toMatch(/\.demo__rb-icon\s*\{[\s\S]*?width: 20px;[\s\S]*?height: 20px;/);
    expect(layoutCss).toMatch(/\.demo__rb--mono\s*\{[\s\S]*?min-width: 30px;/);
    expect(layoutCss).toMatch(
      /\.demo__ribbon-display-option\[aria-checked="true"\]::before\s*\{[\s\S]*?border-bottom: 2px solid var\(--fc-tb-accent\);[\s\S]*?border-left: 2px solid var\(--fc-tb-accent\);[\s\S]*?content: "";[\s\S]*?transform: rotate\(-45deg\);/,
    );
    expect(layoutCss).toMatch(
      /\.demo__ribbon-toggle::before\s*\{[\s\S]*?width: 7px;[\s\S]*?border-top: 1\.6px solid currentColor;[\s\S]*?border-left: 1\.6px solid currentColor;[\s\S]*?content: "";[\s\S]*?rotate\(45deg\);/,
    );
    expect(layoutCss).toMatch(
      /\.demo__ribbon-shell--tabsOnly \.demo__ribbon-toggle::before,[\s\S]*?\.demo__ribbon-shell--autoHide \.demo__ribbon-toggle::before\s*\{[\s\S]*?rotate\(-135deg\);/,
    );
    expect(layoutCss).not.toContain('content: "✓"');
    expect(layoutCss).not.toContain('content: "⌃"');
    expect(layoutCss).not.toContain('content: "⌄"');
    expect(layoutCss).not.toContain(
      'background: var(--fc-tb-ribbon-pressed, var(--fc-tb-accent-soft));',
    );
  });

  it('keeps the titlebar on the neutral Excel 365 desktop surface', () => {
    const tokensCss = readFileSync(join(toolbarBaseDir, 'tokens.css'), 'utf8');
    const titlebarCss = readFileSync(join(toolbarBaseDir, 'header/titlebar.css'), 'utf8');
    const commandbarCss = readFileSync(join(toolbarBaseDir, 'header/commandbar.css'), 'utf8');
    const sidePanelCss = readFileSync(join(toolbarBaseDir, '../panels/side.css'), 'utf8');
    const modalCss = readFileSync(join(toolbarBaseDir, '../panels/modal.css'), 'utf8');
    const backstageCss = readFileSync(join(toolbarBaseDir, 'backstage.css'), 'utf8');

    expect(tokensCss).toContain('--fc-tb-title: #f3f2f1;');
    expect(tokensCss).toContain('--fc-tb-title-fg: #201f1e;');
    expect(titlebarCss).toContain('min-height: 36px;');
    expect(titlebarCss).toContain('background: var(--fc-tb-title);');
    expect(titlebarCss).toContain('background: transparent;');
    expect(titlebarCss).toContain('.demo__search:hover,');
    expect(titlebarCss).toContain('.demo__account .demo__share:first-child');
    expect(titlebarCss).toContain('background: var(--fc-tb-title-strong);');
    expect(commandbarCss).toMatch(
      /\.demo__brand-mark\s*\{[\s\S]*?background: var\(--fc-tb-brand\);[\s\S]*?color: #ffffff;/,
    );
    expect(commandbarCss).toMatch(
      /\.demo__seg-btn:hover\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-hover\);[\s\S]*?color: var\(--fc-tb-fg\);/,
    );
    expect(commandbarCss).toMatch(
      /\.demo__btn:hover:not\(:disabled\)\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-hover\);[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);/,
    );
    expect(commandbarCss).not.toContain('background: var(--fc-tb-accent-soft)');
    expect(sidePanelCss).toMatch(
      /\.demo__card h2\s*\{[\s\S]*?font-size: 12px;[\s\S]*?letter-spacing: 0;[\s\S]*?color: var\(--fc-tb-fg\);/,
    );
    expect(sidePanelCss).toMatch(/\.demo__preset-btn\s*\{[\s\S]*?border-radius: 2px;/);
    expect(sidePanelCss).toMatch(
      /\.demo__preset-btn:hover:not\(:disabled\)\s*\{[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(sidePanelCss).toMatch(
      /\.demo__feat-title\s*\{[\s\S]*?font-size: 11px;[\s\S]*?letter-spacing: 0;[\s\S]*?color: var\(--fc-tb-fg\);/,
    );
    expect(sidePanelCss).toMatch(
      /\.demo__log-cell\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-hover\);[\s\S]*?border-radius: 2px;/,
    );
    expect(sidePanelCss).not.toContain('text-transform: uppercase');
    expect(modalCss).toMatch(
      /\.demo__modal-panel\s*\{[\s\S]*?border-radius: 2px;[\s\S]*?0 12px 32px rgba\(0, 0, 0, 0\.18\),[\s\S]*?0 1px 4px rgba\(0, 0, 0, 0\.14\);/,
    );
    expect(modalCss).toMatch(/\.demo__modal-list li\s*\{[\s\S]*?border-radius: 2px;/);
    expect(modalCss).toMatch(/\.demo__report-item\s*\{[\s\S]*?border-radius: 2px;/);
    expect(modalCss).not.toContain('box-shadow: var(--fc-shadow-16)');
    expect(backstageCss).toMatch(
      /\.demo__backstage-command--active\s*\{[\s\S]*?border-color: var\(--fc-tb-accent\);[\s\S]*?background: var\(--fc-tb-ribbon-pressed, var\(--fc-tb-ribbon-bg\)\);/,
    );
    expect(backstageCss).not.toContain(
      'background: var(--fc-tb-accent-soft, color-mix(in srgb, var(--fc-tb-accent) 12%, transparent));',
    );
  });

  it('keeps ribbon menus on the Excel 365 neutral popup surface', () => {
    const menusCss = readFileSync(join(ribbonStylesDir, 'menus.css'), 'utf8');
    const dropdownsCss = readFileSync(join(ribbonStylesDir, 'dropdowns.css'), 'utf8');
    const colorControlsCss = readFileSync(join(ribbonStylesDir, 'color-controls.css'), 'utf8');

    expect(menusCss).toContain('background: var(--fc-tb-ribbon-bg, #ffffff);');
    expect(menusCss).toContain('border-radius: 2px;');
    expect(menusCss).toMatch(
      /\.app__menu-icon\s*\{[\s\S]*?flex: 0 0 18px;[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
    expect(menusCss).toMatch(
      /\.app__menu-item__icon-spacer\s*\{[\s\S]*?flex: 0 0 18px;[\s\S]*?width: 18px;[\s\S]*?height: 18px;/,
    );
    expect(menusCss).toMatch(
      /\.app__menu-item:hover,[\s\S]*?\.app__submenu-item:focus-visible\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(menusCss).toMatch(
      /\.demo__merge-menu__item:hover,[\s\S]*?\.demo__merge-menu__item:focus-visible\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(menusCss).toMatch(
      /\.app__symbol-tile:hover,[\s\S]*?\.app__symbol-tile:focus-visible\s*\{[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(menusCss).toMatch(
      /\.app__visual-tile:hover,[\s\S]*?\.app__visual-tile:focus-visible\s*\{[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(menusCss).toMatch(
      /\.app__cf-choice:hover,[\s\S]*?\.app__cellstyle-chip:focus-visible\s*\{[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(menusCss).toMatch(
      /\.app__color-swatch--active,[\s\S]*?\.app__color-swatch\[aria-checked="true"\]\s*\{[\s\S]*?border-color: var\(--fc-tb-accent\);[\s\S]*?background: var\(--fc-tb-ribbon-pressed, var\(--fc-tb-ribbon-hover\)\);/,
    );
    expect(menusCss).toMatch(
      /\.app__visual-tile--active,[\s\S]*?\.app__visual-tile\[aria-checked="true"\]\s*\{[\s\S]*?border-color: var\(--fc-tb-accent\);[\s\S]*?background: var\(--fc-tb-ribbon-pressed, var\(--fc-tb-ribbon-hover\)\);/,
    );
    expect(menusCss).toMatch(
      /\.demo__cf-menu__swatch:hover,[\s\S]*?\.demo__cf-menu__iconset:focus-visible\s*\{[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(menusCss).toContain('.app__cf-icon-choice--icons-arrows5 span:nth-child(2)');
    expect(menusCss).toContain('.app__cf-icon-choice--icons-traffic3 span:first-child');
    expect(menusCss).toContain('.app__cf-icon-choice--icons-flags3 span:last-child');
    expect(dropdownsCss).toContain('background: var(--fc-tb-ribbon-bg, var(--fc-tb-input-bg));');
    expect(dropdownsCss).toContain('border-radius: 2px;');
    expect(dropdownsCss).toMatch(
      /\.demo__rb-dd__check\s*\{[\s\S]*?width: 18px;[\s\S]*?min-width: 18px;/,
    );
    expect(dropdownsCss).toContain(
      'background: var(--fc-tb-ribbon-pressed, var(--fc-tb-ribbon-hover));',
    );
    expect(dropdownsCss).toMatch(
      /\.demo__rb-select--font \.demo__rb-dd__list\s*\{[\s\S]*?min-width: 360px;[\s\S]*?padding: 7px 0;[\s\S]*?border-radius: 2px;[\s\S]*?background: var\(--fc-tb-ribbon-bg, #ffffff\);[\s\S]*?box-shadow:/,
    );
    expect(dropdownsCss).toMatch(
      /\.demo__rb-select--font \.demo__rb-dd__section\s*\{[\s\S]*?color: #b5b5b5;/,
    );
    expect(dropdownsCss).toMatch(
      /\.demo__rb-select--font \.demo__rb-dd__opt:hover,[\s\S]*?\.demo__rb-select--font \.demo__rb-dd__opt:focus-visible\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(dropdownsCss).toMatch(
      /\.demo__rb-select--number-format\.demo__rb-dd--open > \.demo__rb-dd__btn,[\s\S]*?\.demo__rb-select--number-format > \.demo__rb-dd__btn:focus-visible\s*\{[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);[\s\S]*?background: var\(--fc-tb-ribbon-hover\);/,
    );
    expect(dropdownsCss).toMatch(
      /\.demo__rb-select--number-format \.demo__rb-dd__list\s*\{[\s\S]*?min-width: 148px;[\s\S]*?padding: 6px 0;[\s\S]*?border-radius: 2px;[\s\S]*?background: var\(--fc-tb-ribbon-bg, #ffffff\);[\s\S]*?box-shadow:/,
    );
    const numberFormatCss = dropdownsCss.slice(
      dropdownsCss.indexOf('.demo__rb-select--number-format .demo__rb-dd__opt'),
      dropdownsCss.indexOf('/* ── Margins picker variant'),
    );
    expect(numberFormatCss).toContain('min-height: 25px;');
    expect(numberFormatCss).toContain('background: #f3f2f1;');
    expect(numberFormatCss).toContain('data:image/svg+xml');
    expect(numberFormatCss).toContain('data-fc-value="percent"');
    expect(numberFormatCss).toContain('%23107c41');
    expect(numberFormatCss).not.toContain('content: "%"');
    expect(numberFormatCss).not.toContain('content: "1/2"');
    expect(numberFormatCss).not.toContain('content: "10²"');
    expect(colorControlsCss).toMatch(
      /\.demo__rb-color__btn:hover,[\s\S]*?\.demo__rb-color--open \.demo__rb-color__btn\s*\{[\s\S]*?background: var\(--fc-tb-ribbon-hover\);[\s\S]*?border-color: var\(--fc-tb-ribbon-line\);/,
    );
    expect(colorControlsCss).toMatch(
      /\.demo__rb-color\s*\{[\s\S]*?width: 30px;[\s\S]*?height: 26px;/,
    );
    expect(colorControlsCss).toMatch(
      /\.demo__rb-color__btn\s*\{[\s\S]*?width: 30px;[\s\S]*?height: 26px;/,
    );
    expect(colorControlsCss).toMatch(
      /\.demo__color-flyout,[\s\S]*?\.demo__merge-menu\s*\{[\s\S]*?border-radius: 3px;[\s\S]*?background: var\(--fc-tb-ribbon-bg, #ffffff\);[\s\S]*?0 8px 18px rgba\(0, 0, 0, 0\.15\)/,
    );
  });

  it('keeps Home dense layout widths large enough to avoid hidden row wraps', () => {
    const groupsCss = readFileSync(join(ribbonStylesDir, 'groups.css'), 'utf8');
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');
    const tileWidth = 74;
    const tileGap = 4;
    const widthFor = (variant: string): number => {
      const match = new RegExp(
        String.raw`\.demo__ribbon-group--${variant} \.demo__ribbon-tools\s*\{\s*width:\s*(\d+)px;`,
        'm',
      ).exec(groupsCss);
      if (!match?.[1]) throw new Error(`missing ${variant} ribbon tool width`);
      return Number(match[1]);
    };

    for (const variant of HOME_TILE_LAYOUT_GROUP_VARIANTS) {
      const commandCount =
        home?.groups.find((group) => group.variant === variant)?.commands.length ?? 0;
      const requiredWidth = commandCount * tileWidth + Math.max(0, commandCount - 1) * tileGap;
      expect(widthFor(variant)).toBeGreaterThanOrEqual(requiredWidth);
    }
    expect(widthFor('cells')).toBeGreaterThanOrEqual(92);
    expect(widthFor('editing')).toBeGreaterThanOrEqual(238);
    expect(widthFor('alignment')).toBeGreaterThanOrEqual(190);
    expect(groupsCss).toMatch(
      /\.demo__ribbon-group--alignment \.demo__rb-icon\s*\{[\s\S]*?width: 22px;[\s\S]*?height: 22px;/,
    );
    expect(groupsCss).toMatch(
      /\.demo__ribbon-group--alignment\.demo__ribbon-group--stacked \.demo__rb--stacked \.demo__rb-icon,[\s\S]*?width: 22px;[\s\S]*?height: 22px;/,
    );
    expect(groupsCss).toContain('.demo__ribbon-group--stacked .demo__ribbon-tools');
    expect(groupsCss).toContain('.demo__ribbon-group--mixed .demo__ribbon-tools');
    expect(groupsCss).toContain('grid-template-rows: repeat(3, 22px);');
    expect(groupsCss).toContain(
      '.demo__ribbon-group--mixed .demo__rb--wide:not(.demo__rb--stacked) .demo__rb-split-chevron',
    );
    expect(groupsCss).toContain(
      '.demo__ribbon-group--stacked .demo__rb--stacked .demo__rb-split-chevron',
    );
    expect(groupsCss).toContain('margin: 0 0 0 auto;');
  });

  it('keeps Home tile-layout groups backed only by wide ribbon commands', () => {
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');
    const nonWideCommands =
      home?.groups
        .filter((group) =>
          HOME_TILE_LAYOUT_GROUP_VARIANTS.includes(
            group.variant as (typeof HOME_TILE_LAYOUT_GROUP_VARIANTS)[number],
          ),
        )
        .flatMap((group) =>
          group.commands
            .filter((command) => command.kind !== 'wide')
            .map((command) => `${group.variant}:${command.id}:${command.kind ?? 'button'}`),
        ) ?? [];

    expect(nonWideCommands).toEqual([]);
  });

  it('keeps Home stacked and mixed groups backed by shared layout metadata', () => {
    const home = buildRibbonModel('en').find((tab) => tab.id === 'home');
    const groups = new Map(home?.groups.map((group) => [group.variant, group]) ?? []);

    expect(HOME_STACKED_LAYOUT_GROUP_VARIANTS).toEqual(['cells']);
    expect(HOME_MIXED_LAYOUT_GROUP_VARIANTS).toEqual(['editing']);
    expect(groups.get('cells')?.commands.map((command) => command.layout)).toEqual([
      'stacked',
      'stacked',
      'stacked',
    ]);
    expect(groups.get('editing')?.commands.map((command) => command.layout ?? 'large')).toEqual([
      'stacked',
      'stacked',
      'stacked',
      'large',
      'large',
    ]);
    expect(groups.get('editing')?.commands.map((command) => command.icon)).toEqual([
      'autosum',
      'fill',
      'clear',
      'sortFilter',
      'find',
    ]);
    expect(
      [...(groups.get('cells')?.commands ?? []), ...(groups.get('editing')?.commands ?? [])].map(
        (command) => command.className ?? '',
      ),
    ).toEqual(['', '', '', '', '', '', '', '']);
  });

  it('can project the Excel 365 standard tab surface without optional add-in tabs', () => {
    const tabs = buildRibbonModel('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS }).map(
      (tab) => tab.id,
    );

    expect(tabs).toEqual([
      'file',
      'home',
      'insert',
      'pageLayout',
      'formulas',
      'data',
      'review',
      'view',
      'help',
    ]);
    expect(tabs).not.toEqual(expect.arrayContaining([...OPTIONAL_RIBBON_TABS]));
  });

  it('keeps Excel 365 command placement for names, duplicates, and links', () => {
    const tabs = new Map(buildRibbonModel('en').map((tab) => [tab.id, tab]));
    const commandIds = (tab: Parameters<typeof ribbonTabCommandIds>[1]): string[] =>
      ribbonTabCommandIds('en', tab);

    expect(commandIds('insert')).not.toEqual(
      expect.arrayContaining(['namedRangesInsert', 'removeDupesInsert', 'linksInsert']),
    );
    expect(commandIds('formulas')).toEqual(expect.arrayContaining(['namedRanges']));
    expect(commandIds('data')).toEqual(expect.arrayContaining(['removeDupes', 'linksData']));
    expect(commandIds('insert')).toEqual(expect.arrayContaining(['hyperlinkInsert']));
    expect(commandIds('pageLayout')).toEqual(
      expect.arrayContaining(['arrangeObjectsPageLayout', 'selectionPanePageLayout']),
    );
    const commands = new Map(ribbonCommands('en').map((command) => [command.id, command]));
    expect(commands.get('formatTableInsert')).toMatchObject({
      label: 'Table',
      title: 'Table',
      icon: 'table',
    });
    expect(commands.get('formatTableHome')).toMatchObject({
      label: 'Format as Table',
      title: 'Format as Table',
    });
    expect(commands.get('pivotTableInsert')).toMatchObject({
      label: 'PivotTable',
      title: 'PivotTable',
      icon: 'pivotTable',
    });
    expect(commands.get('pictureInsert')).toMatchObject({ icon: 'picture' });
    expect(commands.get('shapesInsert')).toMatchObject({ icon: 'shapes' });
    expect(commands.get('screenshotInsert')).toMatchObject({ icon: 'screenshot' });
    expect(commands.get('chartInsert')).toMatchObject({ icon: 'chart' });
    expect(commands.get('hyperlinkInsert')).toMatchObject({ icon: 'link' });
    expect(commands.get('commentInsert')).toMatchObject({ icon: 'commentAdd' });
    expect(commands.get('symbolInsert')).toMatchObject({ icon: 'function' });
    expect(commands.get('pageTheme')).toMatchObject({ icon: 'pageTheme' });
    expect(commands.get('pageSetupAdvanced')).toMatchObject({ icon: 'pageSetup' });
    expect(commands.get('printArea')).toMatchObject({ icon: 'printArea' });
    expect(commands.get('pageBreaks')).toMatchObject({ icon: 'pageBreaks' });
    expect(commands.get('sheetBackground')).toMatchObject({ icon: 'sheetBackground' });
    expect(commands.get('printTitles')).toMatchObject({ icon: 'printTitles' });
    expect(commands.get('filter')).toMatchObject({ icon: 'filter' });
    expect(commands.get('textToColumns')).toMatchObject({ icon: 'textToColumns' });
    expect(commands.get('removeDupes')).toMatchObject({ icon: 'removeDuplicates' });
    expect(commands.get('dataValidation')).toMatchObject({ icon: 'dataValidation' });
    expect(commands.get('outlineGroup')).toMatchObject({ icon: 'outlineGroup' });
    expect(commands.get('outlineUngroup')).toMatchObject({ icon: 'outlineUngroup' });
    expect(commands.get('outlineShowDetail')).toMatchObject({ icon: 'outlineShow' });
    expect(commands.get('outlineHideDetail')).toMatchObject({ icon: 'outlineHide' });
    expect(commands.get('errorChecking')).toMatchObject({ icon: 'errorChecking' });
    expect(commands.get('calcOptions')).toMatchObject({ icon: 'calcOptions' });
    expect(tabs.get('review')?.groups.map((group) => group.title)).toEqual([
      'Proofing',
      'Accessibility',
      'Language',
      'Comments',
      'Find',
      'Protection',
    ]);
  });

  it('has SVG paths for every icon used by the ribbon model', () => {
    const missing = ribbonCommands('en')
      .filter(
        (command) =>
          command.icon && !fluentIconPaths(command.icon) && !excelRibbonIconPaths(command.icon),
      )
      .map((command) => `${command.id}:${command.icon}`);

    expect(missing).toEqual([]);
  });

  it('has Excel-like SVG paths for every visible ribbon model icon', () => {
    const missingExcelIcon = ribbonCommands('en')
      .filter((command) => command.icon && !excelRibbonIconPaths(command.icon))
      .map((command) => `${command.id}:${command.icon}`);

    expect(missingExcelIcon).toEqual([]);
  });

  it('overrides high-value cell formatting icons with Excel-like SVG definitions', () => {
    expect(
      [
        'paste',
        'cut',
        'copy',
        'paint',
        'pasteFormulas',
        'pasteValues',
        'pasteTranspose',
        'pasteSpecial',
        'fontGrow',
        'fontShrink',
        'currency',
        'percent',
        'comma',
        'decDown',
        'decUp',
        'autosum',
        'fill',
        'clear',
        'sortAsc',
        'sortDesc',
        'sortFilter',
        'find',
        'top',
        'middle',
        'bottomAlign',
        'alignLeft',
        'alignCenter',
        'alignRight',
        'textOrientation',
        'wrap',
        'indentDecrease',
        'indentIncrease',
        'merge',
        'table',
        'pivotTable',
        'picture',
        'shapes',
        'screenshot',
        'chart',
        'link',
        'commentAdd',
        'function',
        'pageTheme',
        'pageSetup',
        'printArea',
        'pageBreaks',
        'sheetBackground',
        'printTitles',
        'filter',
        'textToColumns',
        'removeDuplicates',
        'dataValidation',
        'outlineGroup',
        'outlineUngroup',
        'outlineShow',
        'outlineHide',
        'names',
        'trace',
        'dependents',
        'clearArrows',
        'errorChecking',
        'calcOptions',
        'watch',
        'spelling',
        'accessibility',
        'translate',
        'protect',
        'print',
        'freeze',
        'zoom',
        'page',
        'goTo',
        'options',
        'pen',
        'eraser',
        'script',
        'addIn',
        'pdf',
        'borders',
        'fillColor',
        'fontColor',
        'insertRows',
        'insertCols',
        'deleteRows',
        'deleteCols',
        'formatCells',
        'conditional',
        'tableStyle',
        'cellStyles',
      ].filter((name) => !excelRibbonIconPaths(name)),
    ).toEqual([]);
  });

  it('localizes ribbon model command titles for Japanese Office-like surfaces', () => {
    const commands = new Map(ribbonCommands('ja').map((command) => [command.id, command.title]));

    expect(commands.get('numberFormat')).toBe('数値');
    expect(commands.get('paste')).toBe('貼り付け');
    expect(commands.get('bold')).toBe('太字 (Ctrl+B)');
    expect(commands.get('borders')).toBe('罫線');
    const homeFontCommands =
      buildRibbonModel('en')
        .find((tab) => tab.id === 'home')
        ?.groups.find((group) => group.variant === 'font')
        ?.commands.map((command) => command.id) ?? [];
    expect(homeFontCommands.slice(-3)).toEqual(['borders', 'fillColor', 'fontColor']);
    expect(commands.has('borderPreset')).toBe(false);
    expect(commands.has('borderStyle')).toBe(false);
    expect(commands.has('moreBorders')).toBe(false);
    expect(commands.has('drawBorder')).toBe(false);
    expect(commands.has('drawBorderGrid')).toBe(false);
    expect(commands.has('eraseBorder')).toBe(false);
    const modelCommands = ribbonCommands('en');
    const homeAlignmentCommands =
      buildRibbonModel('en')
        .find((tab) => tab.id === 'home')
        ?.groups.find((group) => group.variant === 'alignment')
        ?.commands.map((command) => command.id) ?? [];
    expect(homeAlignmentCommands).toEqual([
      'top',
      'middle',
      'bottomAlign',
      'alignL',
      'alignC',
      'alignR',
      'alignment-row-2',
      'textOrientation',
      'wrap',
      'indentDecrease',
      'indentIncrease',
      'merge',
    ]);
    expect(
      buildRibbonModel('en')
        .find((tab) => tab.id === 'home')
        ?.groups.find((group) => group.variant === 'alignment')
        ?.commands.filter((command) => command.kind !== 'break')
        .map((command) => command.icon),
    ).toEqual([
      'top',
      'middle',
      'bottomAlign',
      'alignLeft',
      'alignCenter',
      'alignRight',
      'textOrientation',
      'wrap',
      'indentDecrease',
      'indentIncrease',
      'merge',
    ]);
    expect(modelCommands.find((command) => command.id === 'wrap')?.kind).toBe('button');
    expect(modelCommands.find((command) => command.id === 'currency')?.kind).toBe('button');
    expect(modelCommands.find((command) => command.id === 'merge')?.kind).toBe('button');
    expect(commands.get('pageSetupAdvanced')).toBe('ページ設定');
    expect(commands.get('sum')).toBe('SUM の引数');
    expect(commands.get('sortAsc')).toBe('昇順で並べ替え');
    expect(commands.get('outlineGroup')).toBe('選択した行または列をグループ化');
    expect(commands.get('deleteCommentReview')).toBe('コメントまたはメモの削除');
    expect(commands.get('viewNormal')).toBe('標準');
    expect(commands.get('viewPageLayout')).toBe('ページ レイアウト');
    expect(commands.get('viewPageBreakPreview')).toBe('改ページ プレビュー');
    expect(commands.get('showFormulasFormula')).toBe('数式');
    expect(commands.get('errorChecking')).toBe('エラー チェック');
    expect(commands.get('evaluateFormula')).toBe('数式の検証');
    expect(commands.get('watchView')).toBe('ウォッチ');
    expect(commands.get('zoom100')).toBe('ズーム 100%');
    expect(commands.get('protect')).toBe('保護');
  });

  it('uses Windows-style shortcut labels for the Japanese Excel 365 desktop baseline', () => {
    const titled = ribbonCommands('ja').filter((command) => command.title);
    const macShortcutTitles = titled
      .filter((command) => command.title.includes('⌘'))
      .map((command) => `${command.id}:${command.title}`);
    const commands = new Map(titled.map((command) => [command.id, command.title]));

    expect(macShortcutTitles).toEqual([]);
    expect(commands.get('bold')).toBe('太字 (Ctrl+B)');
    expect(commands.get('italic')).toBe('斜体 (Ctrl+I)');
    expect(commands.get('underline')).toBe('下線 (Ctrl+U)');
    expect(commands.get('findHome')).toBe('検索と選択 (Ctrl+F)');
    expect(commands.get('findReview')).toBe('検索 (Ctrl+F)');
    expect(commands.get('hyperlinkInsert')).toBe('リンク (Ctrl+K)');
  });

  it('does not expose internal row-break ids as command titles', () => {
    const breaks = buildRibbonModel('ja')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter((command) => command.kind === 'break');

    expect(breaks.map((command) => command.title)).toEqual(breaks.map(() => ''));
  });

  it('keeps Japanese ribbon command titles free of untranslated English phrases', () => {
    const allowed =
      /^(Acrobat|PDF|R1C1|fx|SUM|AVG|AVERAGE|IF|XLOOKUP|CONCAT|TODAY|PMT|ROUND|A-Z|Z-A|\$|%|,|\.0|\.00|B|I|U|S|\d+%)$/;
    const allowedInline =
      /^(?:オートSUM \(Σ\)|.* \(Ctrl\+[A-Z]\)|(?:SUM|AVERAGE|IF|XLOOKUP|CONCAT|TODAY|PMT|ROUND) の引数)$/;
    const untranslated = ribbonCommands('ja')
      .filter((command) => /[A-Za-z]{3,}/.test(command.title))
      .filter((command) => !allowed.test(command.title) && !allowedInline.test(command.title))
      .map((command) => `${command.id}:${command.title}`);

    expect(untranslated).toEqual([]);
  });

  it('keeps shared ribbon menu labels localized for React and Vue wrappers', () => {
    expect(toolbarMenuText('ja').validationSettings).toBe('データの入力規則...');
    expect(toolbarMenuText('en').validationSettings).toBe('Data Validation...');
    expect(toolbarMenuText('ja').errorChecking).toBe('エラー チェック...');
    expect(toolbarMenuText('en').traceError).toBe('Trace Error');
    expect(toolbarMenuText('ja').sortCustom).toBe('ユーザー設定の並べ替え...');
    expect(toolbarMenuText('en').sortCustom).toBe('Custom Sort...');
    expect(toolbarMenuText('ja').symbolGreek).toBe('ギリシャ文字');
    expect(toolbarMenuText('en').symbolMore).toBe('More Symbols...');
    expect(toolbarMenuText('ja').autosumMoreFunctions).toBe('その他の関数...');
    expect(toolbarMenuText('en').autosumMoreFunctions).toBe('More Functions...');
    expect(toolbarMenuText('ja').tableStyleDark).toBe('濃色');
    expect(toolbarMenuText('en').tableStyleDark).toBe('Dark');
    expect(toolbarMenuText('ja').scriptCommandPrompt).toContain(
      'スクリプト コマンドを入力してください',
    );
    expect(toolbarMenuText('ja').autoFitRowHeight).toBe('行の高さの自動調整');
    expect(toolbarMenuText('en').autoFitColWidth).toBe('AutoFit Column Width');
    expect(toolbarMenuText('ja').orientationAngleCounterclockwise).toBe('左回りに回転');
    expect(toolbarMenuText('en').orientationHorizontalText).toBe('Horizontal Text');
    expect(toolbarMenuText('en').scriptCommandPrompt).toBe(
      'Script command: uppercase, lowercase, trim, clear',
    );
    expect(toolbarMenuText('ja').automationRunStatus).toBe(
      'スクリプト · {count} 個のセルを変更しました',
    );
    expect(toolbarMenuText('en').automationRunStatus).toBe('Script · {count} cell(s) changed');
    expect(conditionalMenuText('ja').iconSets).toBe('アイコン セット');
    expect(conditionalMenuText('en').iconSets).toBe('Icon Sets');
    expect(conditionalMenuText('ja').unique).toBe('一意の値...');
    expect(conditionalMenuText('ja').equal).toBe('指定の値に等しい...');
    expect(conditionalMenuText('en').textContains).toBe('Text that Contains...');
    expect(conditionalMenuText('ja').datePrompt).toBe(
      '日付条件: 昨日、今日、明日、過去 7 日間、先週、今週、来週、先月、今月、来月',
    );
    expect(conditionalMenuText('ja').datePrompt).not.toContain('last7');
    expect(conditionalMenuText('en').top10Percent).toBe('Top 10%');
    expect(conditionalMenuText('ja').dataBarSolidGreen).toBe('塗りつぶし (単色)、緑のデータ バー');
    expect(conditionalMenuText('en').dataBarSolidGreen).toBe('Solid Fill, Green Data Bar');
    expect(conditionalMenuText('ja').colorScaleGreenWhiteGreen).toBe(
      '緑 - 白 - 緑のカラー スケール',
    );
    expect(conditionalMenuText('en').colorScaleGreenWhiteGreen).toBe(
      'Green - White - Green Color Scale',
    );
    expect(conditionalMenuText('ja').iconTraffic3).toBe('3 色の信号');
    expect(conditionalMenuText('en').iconStars3).toBe('3 Stars');
    expect(ribbonDisplayText('ja').label).toBe('リボンの表示オプション');
    expect(ribbonDisplayText('en').collapsed).toBe('Show tabs only');
    expect(ribbonDisplayText('en').singleLine).toBe('Single Line Ribbon');
    expect(ribbonDisplayText('en').autoHide).toBe('Auto-hide Ribbon');
    expect(backstageMenuText('ja').back).toBe('戻る');
    expect(backstageMenuText('en').workbookInfo).toBe('Workbook Information');
    expect(pageScaleMenuText('ja').automatic).toBe('自動');
    expect(pageScaleMenuText('en').pages).toBe('pages');
    expect(pageScaleMenuText('ja').custom).toBe('ユーザー設定...');
    expect(pageScaleMenuText('en').invalidScale).toBe('Enter a scale from 10 to 400.');
    expect(viewToggleMenuText('ja').formulaBar).toBe('数式バー');
    expect(viewToggleMenuText('en').gridlines).toBe('Gridlines');
  });
});
