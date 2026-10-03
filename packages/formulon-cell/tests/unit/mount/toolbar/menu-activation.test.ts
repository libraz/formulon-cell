import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { customPivotTableStyleId } from '../../../../src/commands/format-as-table.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import {
  RIBBON_AUDITED_DROPDOWN_COMMANDS,
  RIBBON_AUDITED_GALLERY_COMMANDS,
  RIBBON_AUDITED_PRIMARY_ACTION_SPLIT_COMMANDS,
  RIBBON_AUDITED_SPLIT_TOGGLE_COMMANDS,
  RIBBON_BORDERS_MENU_ID,
  RIBBON_DIALOG_COMMANDS,
  RIBBON_DISABLED_COMMANDS,
  RIBBON_DROPDOWN_COMMANDS,
  RIBBON_DYNAMIC_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FIRST_COMMANDS,
  RIBBON_EXTERNAL_MENU_FOR_COMMAND,
  RIBBON_GALLERY_COMMANDS,
  RIBBON_MENU_FOR_COMMAND,
  RIBBON_PRIMARY_ACTION_COMMANDS,
  RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS,
  RIBBON_PRIMARY_FACE_MENU_COMMANDS,
  RIBBON_SPLIT_BUTTON_COMMANDS,
  RIBBON_SPLIT_TOGGLE_COMMANDS,
  RIBBON_TOGGLE_COMMANDS,
  ribbonActivationCategories,
  ribbonActivationForCommand,
} from '../../../../src/toolbar/ribbon/activation.js';
import {
  RIBBON_DIALOG_OPENERS,
  RIBBON_FUNCTION_ARG_OPENERS,
  RIBBON_HOOK_DIALOG_COMMANDS,
  RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS,
} from '../../../../src/toolbar/ribbon/command-tables.js';
import {
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
} from '../../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import { ribbonActivatableSurfaceCommandIds } from '../../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { dynamicDropdownNoopOverrides, stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('Spreadsheet.mountToolbar', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('closes static fallback ribbon menus on outside mousedown', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    const conditional = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    conditional?.click();
    const conditionalMenu = host.querySelector<HTMLElement>('#menu-conditional');
    expect(conditionalMenu?.hidden).toBe(false);
    expect(conditional?.getAttribute('aria-expanded')).toBe('true');

    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(conditionalMenu?.hidden).toBe(true);
    expect(conditional?.getAttribute('aria-expanded')).toBe('false');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');

    tb.dispose();
  });

  it('closes static fallback ribbon menus on Escape and restores opener focus', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(find);

    tb.dispose();
  });

  it('closes dynamic ribbon dropdowns on Escape and restores opener focus', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    const event = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
    Object.defineProperty(event, 'target', { value: findMenu });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(event)).toBe(true);
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(find);

    tb.dispose();
  });

  it('closes open dynamic ribbon dropdowns when Escape is handled at document scope', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(find);

    tb.dispose();
  });

  it('closes open dynamic ribbon dropdowns when focus moves outside the menu and opener', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const find = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    find?.click();
    const findMenu = host.querySelector<HTMLElement>('#menu-find-select');
    expect(findMenu?.hidden).toBe(false);
    expect(find?.getAttribute('aria-expanded')).toBe('true');

    const firstItem = findMenu?.querySelector<HTMLButtonElement>('button');
    firstItem?.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    expect(findMenu?.hidden).toBe(false);

    const outside = document.createElement('button');
    document.body.appendChild(outside);
    outside.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    expect(findMenu?.hidden).toBe(true);
    expect(find?.getAttribute('aria-expanded')).toBe('false');

    outside.remove();
    tb.dispose();
  });

  it('wires every registered dynamic ribbon dropdown to a rendered button and menu', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    expect(Array.from(tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []).sort()).toEqual(
      Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND).sort(),
    );
    expect(new Set(Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND)).size).toBe(
      Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND).length,
    );

    for (const [command, menuId] of Object.entries(RIBBON_DROPDOWN_MENU_FOR_COMMAND)) {
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      const menu = host.querySelector<HTMLDivElement>(`#${menuId}`);
      expect(button, `${command} button`).toBeTruthy();
      expect(menu, `${command} menu`).toBeTruthy();
      expect(button?.dataset.ribbonMenuId, `${command} menu id metadata`).toBe(menuId);
      expect(button?.getAttribute('aria-haspopup'), `${command} aria-haspopup`).toBe('menu');
      expect(
        button?.querySelector('.fc-tb__rb-split-chevron'),
        `${command} renders dropdown affordance`,
      ).toBeTruthy();
      expect(tb.dropdownsApi?.dynamicDropdownSpecForButton(button as HTMLButtonElement)).toEqual({
        command,
        menuId,
      });
      expect(tb.dropdownsApi?.dynamicDropdownSpecForMenu(menu as HTMLDivElement)).toEqual({
        command,
        menuId,
      });

      tb.dropdownsApi?.openDynamicRibbonDropdown({ command, menuId }, button as HTMLButtonElement);
      expect(menu?.hidden, `${command} opens ${menuId}`).toBe(false);
      expect(button?.getAttribute('aria-expanded'), `${command} aria-expanded`).toBe('true');
      tb.dropdownsApi?.closeDynamicRibbonDropdown({ command, menuId });
      expect(menu?.hidden, `${command} closes ${menuId}`).toBe(true);
    }

    tb.dispose();
  }, 30_000);

  it('opens every dropdown and gallery command menu from primary click', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      onCommand,
    });
    const misses: string[] = [];

    for (const command of RIBBON_DYNAMIC_MENU_FIRST_COMMANDS) {
      const menuId = RIBBON_DROPDOWN_MENU_FOR_COMMAND[command];
      if (!menuId) {
        misses.push(`${command}:menu-map`);
        continue;
      }
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      const menu = host.querySelector<HTMLDivElement>(`#${menuId}`);
      if (!button || !menu) {
        misses.push(`${command}:missing`);
        continue;
      }
      onCommand.mockClear();
      tb.dropdownsApi?.closeDynamicRibbonDropdown({ command, menuId });
      button.click();
      if (menu.hidden) misses.push(`${command}:closed`);
      if (button.getAttribute('aria-expanded') !== 'true') misses.push(`${command}:aria`);
      if (onCommand.mock.calls.length > 0) misses.push(`${command}:command`);
      tb.dropdownsApi?.closeDynamicRibbonDropdown({ command, menuId });
    }

    expect(misses).toEqual([]);
    tb.dispose();
  });

  it('opens every externally owned ribbon menu from primary click', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      onCommand,
    });
    tb.rerender();
    const misses: string[] = [];

    for (const command of RIBBON_EXTERNAL_MENU_FIRST_COMMANDS) {
      const menuId = RIBBON_EXTERNAL_MENU_FOR_COMMAND[command];
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      const menu = host.querySelector<HTMLDivElement>(`#${menuId}`);
      if (!button || !menu) {
        misses.push(`${command}:missing`);
        continue;
      }
      onCommand.mockClear();
      if (!menu.hidden) button.click();
      button.click();
      if (menu.hidden) misses.push(`${command}:closed`);
      if (button.getAttribute('aria-expanded') !== 'true') misses.push(`${command}:aria`);
      if (onCommand.mock.calls.length > 0) misses.push(`${command}:command`);
    }

    expect(misses).toEqual([]);
    tb.dispose();
  });

  it('keeps the shared ribbon activation model aligned with dropdown menus', () => {
    for (const [command, menuId] of Object.entries(RIBBON_MENU_FOR_COMMAND)) {
      expect(RIBBON_SPLIT_BUTTON_COMMANDS.has(command), `${command} renders as menu button`).toBe(
        true,
      );
      expect(ribbonActivationForCommand(command).menuId, `${command} activation menu`).toBe(menuId);
    }

    for (const command of RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS) {
      expect(RIBBON_SPLIT_BUTTON_COMMANDS.has(command), `${command} renders as split`).toBe(true);
      expect(
        RIBBON_DROPDOWN_MENU_FOR_COMMAND[command],
        `${command} has a secondary menu`,
      ).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} activation kind`).toBe(
        'splitPrimary',
      );
    }
    for (const command of RIBBON_SPLIT_TOGGLE_COMMANDS) {
      expect(RIBBON_SPLIT_BUTTON_COMMANDS.has(command), `${command} renders as split toggle`).toBe(
        true,
      );
      expect(
        RIBBON_DROPDOWN_MENU_FOR_COMMAND[command],
        `${command} has a secondary menu`,
      ).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} activation kind`).toBe(
        'splitToggle',
      );
    }

    expect(ribbonActivationForCommand('formatTableHome').kind).toBe('gallery');
    expect(ribbonActivationForCommand('conditional').kind).toBe('gallery');
    expect(ribbonActivationForCommand('dataValidation').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('deleteCommentReview').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('errorChecking').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('protect').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('protectReview').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('script').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('addIn').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('pdf').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('watch').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('watchView').kind).toBe('splitPrimary');
    expect(ribbonActivationForCommand('borders').kind).toBe('dropdown');
    expect(ribbonActivationForCommand('borders').menuId).toBe(RIBBON_BORDERS_MENU_ID);
    expect(RIBBON_EXTERNAL_MENU_FOR_COMMAND.borders).toBe(RIBBON_BORDERS_MENU_ID);
    expect(ribbonActivationForCommand('pageSetup').kind).toBe('dialog');
    expect(ribbonActivationForCommand('printTitles').kind).toBe('dialog');
    expect(ribbonActivationForCommand('sum').kind).toBe('dialog');
    expect(ribbonActivationForCommand('bold').kind).toBe('toggle');
    expect(ribbonActivationForCommand('underline').kind).toBe('splitToggle');
    expect(ribbonActivationForCommand('helpSearch').kind).toBe('disabled');
    expect(RIBBON_PRIMARY_ACTION_COMMANDS.has('print')).toBe(true);
    expect(RIBBON_PRIMARY_ACTION_COMMANDS.has('sheetBackground')).toBe(true);
    expect(RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS.has('pivotTableInsert')).toBe(true);
    expect(ribbonActivationForCommand('formatTableInsert').kind).toBe('dialog');
  });

  it('fixtures the audited menu-backed ribbon activation categories', () => {
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

  it('keeps non-menu ribbon activation sets mutually exclusive', () => {
    const nonMenuSets = ribbonActivationCategories().filter(
      ([kind]) =>
        kind === 'primaryAction' || kind === 'dialog' || kind === 'toggle' || kind === 'disabled',
    );
    const overlaps: string[] = [];
    for (const [index, [leftName, left]] of nonMenuSets.entries()) {
      for (const [rightName, right] of nonMenuSets.slice(index + 1)) {
        for (const command of left) {
          if (right.has(command)) overlaps.push(`${command}:${leftName}/${rightName}`);
        }
      }
    }

    const menuOverlaps = Array.from(RIBBON_PRIMARY_ACTION_COMMANDS)
      .filter((command) => RIBBON_MENU_FOR_COMMAND[command])
      .map((command) => `${command}:primary/menu`);

    expect([...overlaps, ...menuOverlaps]).toEqual([]);
  });

  it('only marks menu-backed commands as gallery or split-primary activations', () => {
    for (const command of RIBBON_GALLERY_COMMANDS) {
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} gallery has menu`).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} gallery activation`).toBe(
        'gallery',
      );
    }
    for (const command of RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS) {
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} split has menu`).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} split activation`).toBe(
        'splitPrimary',
      );
    }
    for (const command of RIBBON_SPLIT_TOGGLE_COMMANDS) {
      expect(RIBBON_MENU_FOR_COMMAND[command], `${command} split toggle has menu`).toBeTruthy();
      expect(ribbonActivationForCommand(command).kind, `${command} split toggle activation`).toBe(
        'splitToggle',
      );
    }
  });

  it('does not leave rendered primary-action commands implicit in the activation model', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const implicit = Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-command]'))
      .filter((el): el is HTMLButtonElement => el instanceof HTMLButtonElement)
      .map((button) => button.dataset.ribbonCommand ?? '')
      .filter((command) => {
        const activation = ribbonActivationForCommand(command);
        return activation.kind === 'primaryAction' && !RIBBON_PRIMARY_ACTION_COMMANDS.has(command);
      })
      .sort();

    expect(implicit).toEqual([]);
    tb.dispose();
  });

  it('projects rendered ribbon button activation metadata from the shared resolver', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const mismatches = Array.from(
      host.querySelectorAll<HTMLButtonElement>('button[data-ribbon-command]'),
    )
      .map((button) => {
        const command = button.dataset.ribbonCommand ?? '';
        const activation = ribbonActivationForCommand(command);
        const actualMenuId = button.dataset.ribbonMenuId;
        if (button.dataset.ribbonActivation !== activation.kind) {
          return `${command}:kind:${button.dataset.ribbonActivation}->${activation.kind}`;
        }
        if (actualMenuId !== activation.menuId) {
          return `${command}:menu:${actualMenuId ?? 'none'}->${activation.menuId ?? 'none'}`;
        }
        return null;
      })
      .filter((mismatch): mismatch is string => mismatch !== null)
      .sort();

    expect(mismatches).toEqual([]);
    tb.dispose();
  });

  it('renders exactly the shared activatable ribbon command surface as buttons', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const renderedButtonCommands = Array.from(
      host.querySelectorAll<HTMLButtonElement>('button[data-ribbon-command]'),
    )
      .map((button) => button.dataset.ribbonCommand ?? '')
      .sort();
    const expectedButtonCommands = ribbonActivatableSurfaceCommandIds().sort();

    expect(renderedButtonCommands).toEqual(expectedButtonCommands);
    tb.dispose();
  });

  it('renders disabled activation commands with disabled accessibility state', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    for (const command of RIBBON_DISABLED_COMMANDS) {
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      expect(button, `${command} button`).toBeTruthy();
      expect(button?.disabled, `${command} disabled`).toBe(true);
      expect(button?.getAttribute('aria-disabled'), `${command} aria-disabled`).toBe('true');
      expect(button?.getAttribute('aria-description'), `${command} aria-description`).toBe(
        'Coming soon',
      );
      expect(button?.dataset.ribbonDisabledReason, `${command} disabled reason`).toBe(
        'Coming soon',
      );
      expect(button?.title, `${command} title`).toContain('Coming soon');
      expect(button?.dataset.ribbonActivation, `${command} activation`).toBe('disabled');
      expect(button?.dataset.ribbonMenuId, `${command} menu`).toBeUndefined();
    }

    tb.dispose();
  });

  it('localizes disabled activation command reasons through shared ribbon text', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      lang: 'ja',
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="helpSearch"]');

    expect(button?.disabled).toBe(true);
    expect(button?.getAttribute('aria-description')).toBe('未実装');
    expect(button?.dataset.ribbonDisabledReason).toBe('未実装');
    expect(button?.title).toContain('未実装');

    tb.dispose();
  });

  it('does not classify ribbon layout row breaks as primary actions', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const rowBreaks = Array.from(host.querySelectorAll<HTMLElement>('.fc-tb__rb-break')).map(
      (el) => el.dataset.ribbonCommand,
    );

    expect(rowBreaks).toEqual(['font-row-2', 'alignment-row-2', 'number-row-2']);
    for (const command of rowBreaks) {
      expect(command).toBeTruthy();
      if (command) expect(RIBBON_PRIMARY_ACTION_COMMANDS.has(command)).toBe(false);
    }

    tb.dispose();
  });

  it('keeps every explicit primary-action command consumable by the shared dispatcher', () => {
    vi.spyOn(sheet.instance, 'openFindReplace').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'print').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'tracePrecedents').mockReturnValue(1);
    vi.spyOn(sheet.instance, 'traceDependents').mockReturnValue(1);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      hooks: {
        automation: {
          allScripts: vi.fn(),
          recordActions: vi.fn(),
        },
        drawing: {
          setInkMode: vi.fn(),
        },
        page: {
          inspect: vi.fn(),
          outline: vi.fn(),
          sheetBackground: vi.fn(),
        },
        protection: {
          allowEditRanges: vi.fn(),
          runWorkbook: vi.fn(),
        },
        review: {
          accessibility: vi.fn(),
          selectComment: vi.fn(),
          spelling: vi.fn(),
          translate: vi.fn(),
        },
        sheetView: {
          deleteActive: vi.fn(),
          save: vi.fn(),
        },
        sortFilter: {
          customSort: vi.fn(),
          removeDuplicates: vi.fn(),
          sort: vi.fn(),
        },
      },
    });
    const missed = Array.from(RIBBON_PRIMARY_ACTION_COMMANDS)
      .filter((command) => tb.applyCommand(command) !== true)
      .sort();

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('keeps every primary-face menu command consumable from its primary face', () => {
    vi.spyOn(sheet.instance, 'openPivotTableDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openDataValidationDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openExternalLinksDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openNamedRangeDialog').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'openPageSetup').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'toggleWatchWindow').mockImplementation(() => undefined);
    vi.spyOn(sheet.instance, 'print').mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      hooks: {
        automation: {
          addInManager: vi.fn(),
          runScript: vi.fn(),
        },
        clipboard: {
          paste: vi.fn(),
        },
        formula: {
          autoSum: vi.fn(),
          errorChecking: vi.fn(),
        },
        insert: {
          createRecommendedChart: vi.fn(),
          insertSymbol: vi.fn(),
        },
        page: {
          pdf: vi.fn(),
          sheetBackground: vi.fn(),
        },
        protection: {
          runSheet: vi.fn(),
        },
        review: {
          deleteComment: vi.fn(),
        },
        sortFilter: {
          splitTextToColumnsCustom: vi.fn(),
        },
      },
    });
    const missed = Array.from(RIBBON_PRIMARY_FACE_MENU_COMMANDS)
      .filter((command) => tb.applyCommand(command) !== true)
      .sort();

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('keeps every toggle-classified command consumable by the shared dispatcher', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
    });
    const missed = Array.from(RIBBON_TOGGLE_COMMANDS)
      .filter((command) => tb.applyCommand(command) !== true)
      .sort();

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('keeps primary dialog commands out of menu-first activation unless explicitly split', () => {
    const overlaps = Object.keys(RIBBON_DIALOG_OPENERS)
      .filter((command) => RIBBON_DROPDOWN_MENU_FOR_COMMAND[command])
      .sort();

    expect(overlaps).toEqual(Array.from(RIBBON_PRIMARY_SPLIT_DIALOG_COMMANDS).sort());
    for (const command of overlaps) {
      expect(RIBBON_PRIMARY_ACTION_SPLIT_COMMANDS.has(command), `${command} is primary split`).toBe(
        true,
      );
      expect(ribbonActivationForCommand(command).kind, `${command} activation`).toBe(
        'splitPrimary',
      );
    }
  });

  it('keeps dialog-classified ribbon commands backed by a shared dispatcher path', () => {
    const backedDialogs = new Set([
      ...Object.keys(RIBBON_DIALOG_OPENERS),
      ...Object.keys(RIBBON_FUNCTION_ARG_OPENERS),
      ...RIBBON_HOOK_DIALOG_COMMANDS,
    ]);

    const missing = Array.from(RIBBON_DIALOG_COMMANDS)
      .filter((command) => !backedDialogs.has(command))
      .sort();
    expect(missing).toEqual([]);
  });

  it('classifies every rendered ribbon menu under a dispatcher or explicit external owner', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const externallyWiredMenuIds = new Set(Object.values(RIBBON_EXTERNAL_MENU_FOR_COMMAND));
    const unowned = Array.from(host.querySelectorAll<HTMLDivElement>('.fc-tb__menu'))
      .map((menu) => menu.id)
      .filter(
        (menuId) =>
          !(tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS.has(menuId) ?? false) &&
          !externallyWiredMenuIds.has(menuId),
      );

    expect(unowned).toEqual([]);
    tb.dispose();
  });

  it('keeps every rendered dynamic dropdown menu item dispatchable', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const missing: string[] = [];

    for (const menuId of tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []) {
      const menu = host.querySelector<HTMLElement>(`#${menuId}`);
      if (!menu) continue;
      for (const button of menu.querySelectorAll<HTMLButtonElement>('button')) {
        const keys = Object.keys(button.dataset);
        if (keys.length === 0) {
          missing.push(`${menuId}:${button.textContent ?? ''}`);
          continue;
        }
        if (!keys.some((key) => DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS.has(key))) {
          missing.push(`${menuId}:${button.textContent ?? ''}:${keys.join(',')}`);
        }
      }
    }

    expect(missing).toEqual([]);
    tb.dispose();
  });

  it('keeps every registered dynamic dropdown handler key represented by default menus', () => {
    mutators.upsertCustomPivotTableStyle(sheet.instance.store, {
      id: customPivotTableStyleId('Dispatch Coverage Pivot'),
      label: 'Dispatch Coverage Pivot',
      style: 'medium',
      color: '#70ad47',
      variant: 'bandedFirstCol',
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const renderedKeys = new Set<string>();

    for (const menuId of tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []) {
      const menu = host.querySelector<HTMLElement>(`#${menuId}`);
      if (!menu) continue;
      for (const element of menu.querySelectorAll<HTMLElement>('*')) {
        for (const key of Object.keys(element.dataset)) renderedKeys.add(key);
      }
    }

    const missing = Array.from(DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS)
      .filter((key) => !renderedKeys.has(key))
      .sort();
    expect(missing).toEqual([]);
    tb.dispose();
  });

  it('keeps rendered ribbon menus out of plain text-only fallback items', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const plainItems: string[] = [];

    const isStructuredMenuButton = (item: HTMLButtonElement): boolean =>
      item.classList.contains('fc-tb__menu-item--iconic') ||
      item.classList.contains('fc-tb__menu-item--preset') ||
      item.classList.contains('fc-tb__cellstyle-chip') ||
      item.classList.contains('fc-tb__tablestyle-swatch') ||
      item.classList.contains('fc-tb__visual-tile') ||
      item.classList.contains('fc-tb__symbol-tile') ||
      item.classList.contains('fc-tb__color-swatch') ||
      item.classList.contains('fc-tb__cf-choice') ||
      item.classList.contains('fc-tb__cf-icon-choice') ||
      item.classList.contains('fc-tb__submenu-item') ||
      item.classList.contains('fc-colorpalette__swatch') ||
      item.classList.contains('fc-colorpalette__action') ||
      !!item.querySelector(
        '.fc-tb__border-preview, .fc-tb__cf-icon, .fc-tb__menu-item__icon-spacer, .fc-tb__text-orientation-preview',
      );

    for (const menu of host.querySelectorAll<HTMLElement>('.fc-tb__menu')) {
      for (const item of menu.querySelectorAll<HTMLButtonElement>('button')) {
        if (!isStructuredMenuButton(item)) {
          plainItems.push(`${menu.id}:${item.textContent?.trim() ?? ''}`);
        }
      }
    }

    expect(plainItems).toEqual([]);
    tb.dispose();
  });

  it('consumes every rendered dynamic dropdown menu item click through the shared dispatcher', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: dynamicDropdownNoopOverrides(),
      helpers: stubHelpers(),
    });
    const missed: string[] = [];

    for (const menuId of tb.dropdownsApi?.DYNAMIC_RIBBON_DROPDOWN_IDS ?? []) {
      const menu = host.querySelector<HTMLElement>(`#${menuId}`);
      if (!menu) continue;
      for (const button of menu.querySelectorAll<HTMLButtonElement>('button')) {
        const event = new MouseEvent('click', { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'target', { value: button });
        if (tb.dropdownsApi?.dynamicRibbonDropdownClick(event) !== true) {
          missed.push(`${menuId}:${button.textContent ?? ''}:${Object.keys(button.dataset)}`);
        }
      }
    }

    expect(missed).toEqual([]);
    tb.dispose();
  });

  it('clamps dynamic dropdowns to the viewport from the shared opener positioning', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 640 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 360 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const freezeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="freeze"]');
    const freezeMenu = host.querySelector<HTMLElement>('#menu-freeze');
    expect(freezeButton).toBeTruthy();
    expect(freezeMenu).toBeTruthy();
    if (freezeButton) {
      Object.defineProperty(freezeButton, 'getBoundingClientRect', {
        value: () =>
          ({
            x: 580,
            y: 310,
            width: 40,
            height: 28,
            top: 310,
            right: 620,
            bottom: 338,
            left: 580,
            toJSON: () => ({}),
          }) as DOMRect,
      });
    }
    if (freezeMenu) {
      Object.defineProperty(freezeMenu, 'offsetWidth', { configurable: true, value: 216 });
      Object.defineProperty(freezeMenu, 'offsetHeight', { configurable: true, value: 420 });
    }

    freezeButton?.click();

    expect(freezeMenu?.style.position).toBe('fixed');
    expect(freezeMenu?.style.left).toBe('416px');
    expect(freezeMenu?.style.top).toBe('8px');
    expect(freezeMenu?.style.maxHeight).toBe('299px');
    expect(freezeMenu?.style.overflowY).toBe('auto');
    expect(freezeMenu?.style.overscrollBehavior).toBe('contain');

    tb.dispose();
  });
});
