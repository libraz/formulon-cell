import { describe, expect, it } from 'vitest';
import { menuSources, sourcesOutsidePrimitives } from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
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
          source.includes("'fc-tb__menu-icon'"),
      )
      .map(({ name }) => name);

    expect(directIconicRows).toEqual([]);
  });

  it('keeps raw menu button creation centralized in createMenuButton', () => {
    const directButtons = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes("document.createElement('button')"))
      .map(({ name }) => name);

    expect(directButtons).toEqual([]);
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
});
