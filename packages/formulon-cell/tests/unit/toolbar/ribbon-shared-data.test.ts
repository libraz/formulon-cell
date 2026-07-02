import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  RIBBON_AUDITED_DROPDOWN_COMMANDS,
  RIBBON_AUDITED_GALLERY_COMMANDS,
  RIBBON_AUDITED_PRIMARY_ACTION_SPLIT_COMMANDS,
  RIBBON_AUDITED_SPLIT_TOGGLE_COMMANDS,
  RIBBON_BORDERS_MENU_ID,
  RIBBON_DIALOG_COMMANDS,
  RIBBON_DROPDOWN_COMMANDS,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
  RIBBON_DYNAMIC_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FOR_COMMAND,
  RIBBON_GALLERY_COMMANDS,
  RIBBON_INTENTIONAL_NON_RENDERED_COMMANDS,
  RIBBON_MENU_FACTORY_FOR_COMMAND,
  RIBBON_MENU_FACTORY_KEYS,
  RIBBON_MENU_FIRST_COMMANDS,
  RIBBON_MENU_FOR_COMMAND,
  RIBBON_PRIMARY_ACTION_COMMANDS,
  RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS,
  RIBBON_PRIMARY_FACE_MENU_COMMANDS,
  RIBBON_SPLIT_TOGGLE_COMMANDS,
  RIBBON_TOGGLE_COMMANDS,
  ribbonActivationCategories,
  ribbonActivationCommandIds,
  ribbonActivationEntries,
  ribbonActivationEntriesForCommands,
  ribbonActivationForCommand,
} from '../../../src/toolbar/ribbon/activation.js';
import {
  RIBBON_BORDER_DRAW_MODES,
  RIBBON_DIALOG_OPENERS,
  RIBBON_FORMAT_MUTATORS,
  RIBBON_FUNCTION_ARG_OPENERS,
  RIBBON_HOOK_DIALOG_COMMANDS,
  RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS,
  RIBBON_VIEW_MODES,
  RIBBON_ZOOM_PRESETS,
} from '../../../src/toolbar/ribbon/command-tables.js';
import { RIBBON_ACTIVE_COMMANDS } from '../../../src/toolbar/ribbon-active-state.js';
import {
  EXCEL365_STANDARD_RIBBON_TABS,
  OPTIONAL_RIBBON_TABS,
  ribbonActivatableSurfaceCommandIds,
  ribbonActivatableSurfaceCommands,
  ribbonSurfaceCommandIds,
  ribbonSurfaceCommands,
} from '../../../src/toolbar/ribbon-model.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const source = (path: string): string => readFileSync(join(root, path), 'utf8');

const sourceFilesUnder = (path: string): string[] => {
  const absolutePath = join(root, path);
  try {
    if (!statSync(absolutePath).isDirectory()) return [path];
  } catch {
    return [];
  }

  const files: string[] = [];
  const visit = (directory: string): void => {
    for (const entry of readdirSync(directory)) {
      const absoluteEntry = join(directory, entry);
      const stats = statSync(absoluteEntry);
      if (stats.isDirectory()) {
        visit(absoluteEntry);
      } else {
        files.push(relative(root, absoluteEntry));
      }
    }
  };
  visit(absolutePath);
  return files.sort();
};

const renderedCommandIds = (): Set<string> =>
  new Set(ribbonSurfaceCommandIds({ tabs: EXCEL365_STANDARD_RIBBON_TABS }));

const activationManifestByCommand = (): Map<
  string,
  ReturnType<typeof ribbonActivationEntries>[number]
> => new Map(ribbonActivationEntries().map((entry) => [entry.command, entry]));

const objectKeysFromSource = (contents: string, objectName: string): string[] => {
  const block = contents.match(
    new RegExp(
      `const ${objectName}: Readonly<Record<string, (?:string|number)>> = \\{([\\s\\S]*?)\\n\\};`,
    ),
  )?.[1];
  expect(block, objectName).toBeDefined();
  return Array.from(block?.matchAll(/^\s*([A-Za-z0-9]+):/gm) ?? [])
    .flatMap((match) => (match[1] ? [match[1]] : []))
    .sort();
};

const switchCaseLabelsFromSource = (contents: string): string[] =>
  Array.from(contents.matchAll(/case '([^']+)':/g))
    .flatMap((match) => (match[1] ? [match[1]] : []))
    .sort();

describe('toolbar/ribbon shared data', () => {
  it('keeps font dropdown grouping data shared with font availability', () => {
    const selectColorSource = source('src/toolbar/ribbon/select-color.ts');

    expect(selectColorSource).toContain("from './font-availability.js'");
    expect(selectColorSource).toContain('THEME_FONT_VALUES');
    expect(selectColorSource).toContain('RECENT_FONT_VALUES');
    expect(selectColorSource).toContain('FONT_SUBMENU_FAMILIES');
    expect(selectColorSource).not.toContain('const themeFontValues = new Set');
    expect(selectColorSource).not.toContain('const recentFontValues = new Set');
    expect(selectColorSource).not.toContain('const commonFontValues = new Set');
    expect(selectColorSource).not.toContain('const fontSubmenuFamilies = new Set');
  });

  it('keeps renderer menu factories aligned with activation menu commands', () => {
    const renderRibbonSource = source('src/toolbar/ribbon/render-ribbon.ts');
    const factoryCommands = Object.keys(RIBBON_MENU_FACTORY_FOR_COMMAND).sort();

    expect(factoryCommands).toEqual(Object.keys(RIBBON_MENU_FOR_COMMAND).sort());
    expect(renderRibbonSource).toContain('RIBBON_MENU_FACTORY_FOR_COMMAND');
    expect(renderRibbonSource).not.toContain('const MENU_ROUTES');
  });

  it('keeps top-level disabled reason projection on the shared helper', () => {
    const renderRibbonSource = source('src/toolbar/ribbon/render-ribbon.ts');

    expect(renderRibbonSource).toContain('projectDisabledState(button, disabled, disabledReason');
    expect(renderRibbonSource).not.toContain("b.setAttribute('aria-description'");
    expect(renderRibbonSource).not.toContain('b.dataset.ribbonDisabledReason');
  });

  it('keeps ribbon renderer button DOM centralized in local helpers', () => {
    const renderRibbonSource = source('src/toolbar/ribbon/render-ribbon.ts');
    const buttonSource = source('src/toolbar/ribbon/button.ts');

    expect(renderRibbonSource).toContain("import { createRibbonButton } from './button.js'");
    expect(renderRibbonSource).toContain('const createRibbonTabButton');
    expect(renderRibbonSource).toContain('const createRibbonCommandButton');
    expect(renderRibbonSource).toContain('const createRibbonDisplayToggleButton');
    expect(renderRibbonSource).toContain('const createRibbonDisplayOptionButton');
    expect(renderRibbonSource).toContain('tabs.appendChild(createRibbonTabButton(');
    expect(renderRibbonSource).toContain('const b = createRibbonCommandButton(');
    expect(renderRibbonSource).toContain('const toggle = createRibbonDisplayToggleButton(');
    expect(renderRibbonSource).toContain('const item = createRibbonDisplayOptionButton(');
    expect(renderRibbonSource).not.toContain("const btn = document.createElement('button')");
    expect(renderRibbonSource).not.toContain("const b = document.createElement('button')");
    expect(renderRibbonSource).not.toContain("const toggle = document.createElement('button')");
    expect(renderRibbonSource).not.toContain("const item = document.createElement('button')");
    expect(renderRibbonSource).not.toContain("document.createElement('button')");
    expect(buttonSource.match(/document\.createElement\('button'\)/g) ?? []).toHaveLength(1);
  });

  it('keeps menu factory keys shared between activation metadata and renderer slots', () => {
    const renderRibbonSource = source('src/toolbar/ribbon/render-ribbon.ts');
    const menusBlock = renderRibbonSource.match(
      /export interface RibbonMenus \{([\s\S]*?)\n\}/,
    )?.[1];
    expect(menusBlock, 'RibbonMenus block').toBeDefined();
    const menuSlots = Array.from(menusBlock?.matchAll(/^\s*([A-Za-z0-9]+)\?:/gm) ?? [])
      .flatMap((match) => (match[1] ? [match[1]] : []))
      .sort();

    expect(Array.from(RIBBON_MENU_FACTORY_KEYS).sort()).toEqual(menuSlots);
    for (const key of Object.values(RIBBON_MENU_FACTORY_FOR_COMMAND)) {
      expect(RIBBON_MENU_FACTORY_KEYS).toContain(key);
    }
  });

  it('keeps hook-backed dialog commands declared in shared command tables', () => {
    const commandTablesSource = source('src/toolbar/ribbon/command-tables.ts');
    const applyCommandSource = source('src/toolbar/ribbon/apply-ribbon-command.ts');

    expect(Array.from(RIBBON_HOOK_DIALOG_COMMANDS).sort()).toEqual([
      'formatTableInsert',
      'zoomDialog',
    ]);
    for (const command of RIBBON_HOOK_DIALOG_COMMANDS) {
      expect(RIBBON_DIALOG_COMMANDS.has(command), `${command} activation`).toBe(true);
      expect(RIBBON_DIALOG_OPENERS[command], `${command} instance opener`).toBeUndefined();
      expect(RIBBON_FUNCTION_ARG_OPENERS[command], `${command} function opener`).toBeUndefined();
      expect(applyCommandSource).toContain(`case '${command}':`);
    }
    expect(commandTablesSource).toContain('RIBBON_HOOK_DIALOG_COMMANDS');
  });

  it('keeps function-argument dialog commands only in the function opener table', () => {
    for (const command of Object.keys(RIBBON_FUNCTION_ARG_OPENERS)) {
      expect(RIBBON_DIALOG_COMMANDS.has(command), `${command} activation`).toBe(true);
      expect(RIBBON_DIALOG_OPENERS[command], `${command} instance opener`).toBeUndefined();
      expect(RIBBON_HOOK_DIALOG_COMMANDS.has(command), `${command} hook opener`).toBe(false);
    }
  });

  it('keeps instance dialog openers classified as dialog or explicit split-primary dialogs', () => {
    for (const command of Object.keys(RIBBON_DIALOG_OPENERS)) {
      const isDialog = RIBBON_DIALOG_COMMANDS.has(command);
      const isSplitDialog =
        RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS.has(command) &&
        RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS.has(command);

      expect(isDialog || isSplitDialog, `${command} activation`).toBe(true);
      expect(RIBBON_FUNCTION_ARG_OPENERS[command], `${command} function opener`).toBeUndefined();
      expect(RIBBON_HOOK_DIALOG_COMMANDS.has(command), `${command} hook opener`).toBe(false);
    }
  });

  it('keeps primary split dialog commands declared once in shared command tables', () => {
    const commandTablesSource = source('src/toolbar/ribbon/command-tables.ts');
    const menuDialogOpeners = Object.keys(RIBBON_DIALOG_OPENERS)
      .filter((command) => RIBBON_DROPDOWN_MENU_FOR_COMMAND[command])
      .sort();

    expect(Array.from(RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS).sort()).toEqual(menuDialogOpeners);
    for (const command of RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS) {
      expect(RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS.has(command), `${command} split activation`).toBe(
        true,
      );
      expect(RIBBON_DIALOG_OPENERS[command], `${command} dialog opener`).toBeTruthy();
      expect(RIBBON_DROPDOWN_MENU_FOR_COMMAND[command], `${command} secondary menu`).toBeTruthy();
    }
    expect(commandTablesSource).toContain('RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS');
  });

  it('keeps declarative dispatcher tables aligned with activation categories', () => {
    for (const command of Object.keys(RIBBON_ZOOM_PRESETS)) {
      expect(RIBBON_PRIMARY_ACTION_COMMANDS.has(command), `${command} zoom activation`).toBe(true);
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} menu route`).toBeUndefined();
    }
    for (const command of Object.keys(RIBBON_VIEW_MODES)) {
      expect(RIBBON_PRIMARY_ACTION_COMMANDS.has(command), `${command} view activation`).toBe(true);
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} menu route`).toBeUndefined();
    }
    for (const command of Object.keys(RIBBON_BORDER_DRAW_MODES)) {
      expect(RIBBON_TOGGLE_COMMANDS.has(command), `${command} border draw activation`).toBe(true);
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} menu route`).toBeUndefined();
    }
    for (const command of RIBBON_SPLIT_TOGGLE_COMMANDS) {
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} split-toggle menu`).toBeTruthy();
      expect(RIBBON_FORMAT_MUTATORS[command], `${command} split-toggle mutator`).toBeTruthy();
    }
  });

  it('keeps executable activation commands wired to dispatcher entrypoints', () => {
    const applyCommandSource = source('src/toolbar/ribbon/apply-ribbon-command.ts');
    const dispatcherCommands = new Set([
      ...switchCaseLabelsFromSource(applyCommandSource),
      ...Object.keys(RIBBON_DIALOG_OPENERS),
      ...Object.keys(RIBBON_FUNCTION_ARG_OPENERS),
      ...Object.keys(RIBBON_FORMAT_MUTATORS),
      ...Object.keys(RIBBON_ZOOM_PRESETS),
      ...Object.keys(RIBBON_VIEW_MODES),
      ...Object.keys(RIBBON_BORDER_DRAW_MODES),
    ]);
    const executableActivationKinds = new Set([
      'primaryAction',
      'splitPrimary',
      'splitToggle',
      'dialog',
      'toggle',
    ]);
    const missing = ribbonActivationEntries()
      .filter((entry) => executableActivationKinds.has(entry.kind))
      .map((entry) => entry.command)
      .filter((command) => !dispatcherCommands.has(command))
      .sort();

    expect(missing).toEqual([]);
  });

  it('keeps split-toggle commands projected through the shared active-state table', () => {
    const missing = Array.from(RIBBON_SPLIT_TOGGLE_COMMANDS)
      .filter((command) => !RIBBON_ACTIVE_COMMANDS.has(command))
      .sort();

    expect(missing).toEqual([]);
  });

  it('keeps mount toolbar using the shared active-state command map', () => {
    const toolbarSource = source('src/mount/toolbar.ts');

    expect(toolbarSource).toContain(
      "RIBBON_ACTIVE_COMMANDS } from '../toolbar/ribbon-active-state.js'",
    );
    expect(toolbarSource).not.toContain('const RIBBON_ACTIVE_COMMANDS');
  });

  it('keeps rendered format mutator commands in executable activation categories', () => {
    const renderedCommands = new Set(ribbonSurfaceCommandIds());
    const executableKinds = new Set([
      'primaryAction',
      'toggle',
      'splitPrimary',
      'splitToggle',
      'dropdown',
    ]);
    const invalid = Object.keys(RIBBON_FORMAT_MUTATORS)
      .filter((command) => renderedCommands.has(command))
      .filter((command) => !executableKinds.has(ribbonActivationForCommand(command).kind))
      .sort();

    expect(invalid).toEqual([]);
  });

  it('keeps rendered button commands explicitly classified in the activation model', () => {
    const manifest = activationManifestByCommand();
    const implicit = ribbonActivatableSurfaceCommands()
      .map((command) => command.id)
      .filter((command) => !manifest.has(command))
      .sort();

    expect(implicit).toEqual([]);
  });

  it('keeps non-rendered activation commands intentional and documented', () => {
    const renderedCommands = new Set(ribbonSurfaceCommandIds());
    const nonRendered = ribbonActivationEntries()
      .map((entry) => entry.command)
      .filter((command) => !renderedCommands.has(command))
      .sort();

    expect(nonRendered).toEqual([...RIBBON_INTENTIONAL_NON_RENDERED_COMMANDS].sort());
  });

  it('exposes a reusable activation manifest for every explicit command', () => {
    const entries = ribbonActivationEntries();
    const commandIds = entries.map((entry) => entry.command);
    const mismatches = entries
      .map((entry) => {
        const actual = ribbonActivationForCommand(entry.command);
        if (actual.kind !== entry.kind) {
          return `${entry.command}:kind:${entry.kind}->${actual.kind}`;
        }
        if (actual.menuId !== entry.menuId) {
          return `${entry.command}:menu:${entry.menuId ?? 'none'}->${actual.menuId ?? 'none'}`;
        }
        return null;
      })
      .filter((mismatch): mismatch is string => mismatch !== null);

    expect(commandIds).toEqual(ribbonActivationCommandIds());
    expect(new Set(commandIds).size).toBe(commandIds.length);
    expect(mismatches).toEqual([]);
  });

  it('exposes reusable activation manifests for selected command surfaces', () => {
    const standardCommandIds = ribbonActivatableSurfaceCommandIds({
      tabs: EXCEL365_STANDARD_RIBBON_TABS,
    });
    const optionalCommandIds = ribbonActivatableSurfaceCommandIds({ tabs: OPTIONAL_RIBBON_TABS });
    const standardEntries = ribbonActivationEntriesForCommands(standardCommandIds);
    const optionalEntries = ribbonActivationEntriesForCommands(optionalCommandIds);

    expect(standardEntries.map((entry) => entry.command)).toEqual(
      Array.from(new Set(standardCommandIds)).sort(),
    );
    expect(optionalEntries.map((entry) => entry.command)).toEqual(
      Array.from(new Set(optionalCommandIds)).sort(),
    );
    expect(
      standardEntries.filter((entry) => entry.kind === 'disabled').map((entry) => entry.command),
    ).toEqual(['helpSearch']);
    expect(optionalEntries.filter((entry) => entry.kind === 'disabled')).toEqual([]);
  });

  it('keeps effective activation categories mutually exclusive', () => {
    const categories = ribbonActivationCategories();
    const overlaps: string[] = [];

    for (const [index, [leftName, left]] of categories.entries()) {
      for (const [rightName, right] of categories.slice(index + 1)) {
        for (const command of left) {
          if (right.has(command)) overlaps.push(`${command}:${leftName}/${rightName}`);
        }
      }
    }

    expect(overlaps).toEqual([]);
  });

  it('keeps primary-face menu commands derived from split-primary and split-toggle sets', () => {
    expect(Array.from(RIBBON_PRIMARY_FACE_MENU_COMMANDS).sort()).toEqual(
      [...RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS, ...RIBBON_SPLIT_TOGGLE_COMMANDS].sort(),
    );
  });

  it('keeps menu-first commands derived from menu-backed commands without primary faces', () => {
    expect(Array.from(RIBBON_MENU_FIRST_COMMANDS).sort()).toEqual(
      Object.keys(RIBBON_MENU_FOR_COMMAND)
        .filter((command) => !RIBBON_PRIMARY_FACE_MENU_COMMANDS.has(command))
        .sort(),
    );
  });

  it('keeps menu-first owner sets aligned with dynamic and external menu maps', () => {
    expect(Array.from(RIBBON_DYNAMIC_MENU_FIRST_COMMANDS).sort()).toEqual(
      Object.keys(RIBBON_DROPDOWN_MENU_FOR_COMMAND)
        .filter((command) => RIBBON_MENU_FIRST_COMMANDS.has(command))
        .sort(),
    );
    expect(Array.from(RIBBON_EXTERNAL_MENU_FIRST_COMMANDS).sort()).toEqual(
      Object.keys(RIBBON_EXTERNAL_MENU_FOR_COMMAND)
        .filter((command) => RIBBON_MENU_FIRST_COMMANDS.has(command))
        .sort(),
    );
    expect(
      [...RIBBON_DYNAMIC_MENU_FIRST_COMMANDS, ...RIBBON_EXTERNAL_MENU_FIRST_COMMANDS].sort(),
    ).toEqual(Array.from(RIBBON_MENU_FIRST_COMMANDS).sort());
    expect(RIBBON_EXTERNAL_MENU_FOR_COMMAND.borders).toBe(RIBBON_BORDERS_MENU_ID);
  });

  it('keeps the main Borders menu id owned by the activation model', () => {
    const activationSource = source('src/toolbar/ribbon/activation.ts');
    const ownedLiteral = /export const RIBBON_BORDERS_MENU_ID = 'menu-borders';/;
    const exactMenuLiteral = /['"]menu-borders['"]/;
    const exactMenuSelectorLiteral = /['"]#menu-borders['"]/;
    const consumers = [
      'src/toolbar/ribbon/menus/borders.ts',
      'src/toolbar/ribbon/border-menu.ts',
      'src/mount/toolbar.ts',
      'tests/unit/mount/toolbar.test.ts',
    ];

    expect(activationSource).toMatch(ownedLiteral);
    expect(RIBBON_BORDERS_MENU_ID).toBe('menu-borders');
    for (const path of consumers) {
      const contents = source(path);
      expect(contents, `${path} menu id literal`).not.toMatch(exactMenuLiteral);
      expect(contents, `${path} menu selector literal`).not.toMatch(exactMenuSelectorLiteral);
    }
  });

  it('keeps dropdown commands derived from menu-first commands without galleries', () => {
    expect(Array.from(RIBBON_DROPDOWN_COMMANDS).sort()).toEqual(
      Array.from(RIBBON_MENU_FIRST_COMMANDS)
        .filter((command) => !RIBBON_GALLERY_COMMANDS.has(command))
        .sort(),
    );
  });

  it('keeps audited menu-backed activation fixtures as shared source data', () => {
    expect(Array.from(RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_PRIMARY_ACTION_SPLIT_COMMANDS].sort(),
    );
    expect(Array.from(RIBBON_SPLIT_TOGGLE_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_SPLIT_TOGGLE_COMMANDS].sort(),
    );
    expect(Array.from(RIBBON_GALLERY_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_GALLERY_COMMANDS].sort(),
    );
    expect(Array.from(RIBBON_DROPDOWN_COMMANDS).sort()).toEqual(
      [...RIBBON_AUDITED_DROPDOWN_COMMANDS].sort(),
    );
  });

  it('keeps effective activation categories aligned with activation resolution', () => {
    const mismatches: string[] = [];

    for (const [kind, commands] of ribbonActivationCategories()) {
      for (const command of commands) {
        const actual = ribbonActivationForCommand(command).kind;
        if (actual !== kind) mismatches.push(`${command}:${kind}->${actual}`);
      }
    }

    expect(mismatches).toEqual([]);
  });

  it('does not treat unknown ribbon commands as implicit primary actions', () => {
    expect(ribbonActivationForCommand('__unknown_command__')).toEqual({ kind: 'disabled' });
  });

  it('keeps disabled ribbon model commands aligned with disabled activation commands', () => {
    const disabledModelCommands = ribbonSurfaceCommands()
      .filter((command) => command.disabled)
      .map((command) => command.id)
      .sort();
    const disabledActivationCommands = ribbonActivationEntries()
      .filter((entry) => entry.kind === 'disabled')
      .map((entry) => entry.command)
      .sort();

    expect(disabledModelCommands).toEqual(disabledActivationCommands);
  });

  it('keeps the standard Excel 365 command surface separate from optional tabs', () => {
    const standardEntries = ribbonActivationEntriesForCommands(
      ribbonActivatableSurfaceCommandIds({ tabs: EXCEL365_STANDARD_RIBBON_TABS }),
    );
    const optionalEntries = ribbonActivationEntriesForCommands(
      ribbonActivatableSurfaceCommandIds({ tabs: OPTIONAL_RIBBON_TABS }),
    );
    const standardCommands = new Set(standardEntries.map((entry) => entry.command));
    const optionalCommands = optionalEntries.map((entry) => entry.command);

    expect(optionalCommands).toEqual([
      'addIn',
      'allScripts',
      'drawErase',
      'drawGrid',
      'drawPen',
      'pdf',
      'recordActions',
      'script',
    ]);
    expect(optionalCommands.filter((command) => standardCommands.has(command))).toEqual([]);
    expect(optionalEntries.filter((entry) => entry.kind === 'disabled')).toEqual([]);
  });

  it('keeps search alias and boost command keys aligned with rendered ribbon commands', () => {
    const searchSource = source('src/toolbar/ribbon/search-index.ts');
    const modelCommands = renderedCommandIds();

    for (const command of objectKeysFromSource(searchSource, 'COMMAND_SEARCH_ALIASES')) {
      expect(modelCommands.has(command), `${command} alias command`).toBe(true);
    }
    for (const command of objectKeysFromSource(searchSource, 'COMMAND_SEARCH_BOOSTS')) {
      expect(modelCommands.has(command), `${command} boost command`).toBe(true);
    }
  });

  it('keeps Search/Tell me using the shared activatable ribbon predicate', () => {
    const searchSource = source('src/toolbar/ribbon/search-index.ts');

    expect(searchSource).toContain('isRibbonActivatableCommand');
    expect(searchSource).not.toContain('ACTIVATABLE_SEARCH_COMMAND_KINDS');
    expect(searchSource).not.toContain("['select', 'color']");
    expect(searchSource).not.toContain('"select", "color"');
  });

  it('keeps ribbon surface helpers exported from the public entrypoint', () => {
    const indexSource = source('src/index.ts');

    for (const symbol of [
      'isRibbonActivatableCommand',
      'ribbonActivatableCommandIds',
      'ribbonActivatableCommands',
      'ribbonActivatableSurfaceCommandIds',
      'ribbonActivatableSurfaceCommands',
      'ribbonSurfaceCommandIds',
      'ribbonSurfaceCommands',
    ]) {
      expect(indexSource, symbol).toContain(symbol);
    }
  });

  it('keeps framework toolbar wrappers delegated to the core toolbar implementation', () => {
    const reactToolbarFiles = sourceFilesUnder('../formulon-cell-react/src/toolbar');
    const vueToolbarFiles = sourceFilesUnder('../formulon-cell-vue/src/toolbar');
    const reactToolbarSource = source('../formulon-cell-react/src/SpreadsheetToolbar.tsx');
    const vueToolbarSource = source('../formulon-cell-vue/src/SpreadsheetToolbar.vue');
    const vueToolbarDts = source('../formulon-cell-vue/src/SpreadsheetToolbar.vue.d.ts');
    const vueToolbarExports = source('../formulon-cell-vue/src/toolbar.ts');
    const reactIndexSource = source('../formulon-cell-react/src/index.ts');
    const vueIndexSource = source('../formulon-cell-vue/src/index.ts');

    expect(reactToolbarFiles).toEqual(['../formulon-cell-react/src/toolbar/model.ts']);
    expect(vueToolbarFiles).toEqual([]);
    expect(reactToolbarSource).toContain('Spreadsheet.mountToolbar');
    expect(vueToolbarSource).toContain('Spreadsheet.mountToolbar');
    expect(reactToolbarSource).toContain('callbacksRef.current.onError?.(error)');
    expect(vueToolbarSource).toContain("emit('error', error)");
    expect(reactToolbarSource).toContain('export const Toolbar = SpreadsheetToolbar');
    expect(vueToolbarSource).toContain('export const Toolbar = SpreadsheetToolbar');
    expect(vueToolbarDts).toContain('export const Toolbar: typeof SpreadsheetToolbar');
    for (const symbol of ['DynamicDropdownsCtx', 'RibbonTab', 'ToolbarInstance']) {
      expect(vueToolbarExports).toContain(symbol);
      expect(reactIndexSource).toContain(symbol);
      expect(vueIndexSource).toContain(symbol);
    }
    for (const symbol of [
      'EXCEL365_STANDARD_RIBBON_TABS',
      'OPTIONAL_RIBBON_TABS',
      'RIBBON_TAB_LABELS',
      'RIBBON_TABS',
    ]) {
      expect(reactIndexSource).toContain(symbol);
      expect(vueIndexSource).toContain(symbol);
    }
    expect(vueIndexSource).toContain('SpreadsheetToolbarProps');
  });
});
