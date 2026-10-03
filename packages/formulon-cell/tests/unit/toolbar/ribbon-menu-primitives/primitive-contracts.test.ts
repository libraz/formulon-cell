import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
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
} from '../../../../src/toolbar/ribbon/menus/general.js';
import { readCssWithImports } from '../../../helpers/css.js';
import { root } from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
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

    const menusCss = readCssWithImports(join(root, 'src/styles/toolbar/ribbon/menus.css'));
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
