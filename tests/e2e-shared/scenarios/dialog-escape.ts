import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type Platform = 'default' | 'mac';
type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** The workbook state a dialog needs before it can open. */
export type DialogEscapeFixture = 'blank' | 'copied' | 'filled' | 'misspelt' | 'chart' | 'pivot';

export const DIALOG_ESCAPE_FIXTURES: readonly DialogEscapeFixture[] = [
  'blank',
  'copied',
  'filled',
  'misspelt',
  'chart',
  'pivot',
];

type FixtureLists = Record<DialogEscapeFixture, readonly string[]>;

/** Modal dialogs (aria-modal): Escape closes them and hands focus back to the control that opened
 *  them, or to the sheet when that control is gone; never to the page body. */
const MODALS: Record<Platform, FixtureLists> = {
  default: {
    blank: [
      'autoCorrect',
      'cellStylesGallery',
      'cfRules',
      'conditional',
      'dataValidation',
      'defineName',
      'evaluateFormula',
      'format',
      'goTo',
      'goToSpecial',
      'hyperlink',
      'iterative',
      'namedRange',
      'pageSetup',
      'pivotTable',
      'accessibilityChecker',
      'externalLinks',
      'functionArguments',
    ],
    copied: ['pasteSpecial', 'insertCopiedCells'],
    filled: [],
    misspelt: ['spelling'],
    chart: [],
    pivot: ['pivotFieldSettings', 'pivotOptions'],
  },
  mac: {
    blank: [
      'autoCorrect',
      'cellStylesGallery',
      'cfRules',
      'conditional',
      'dataValidation',
      'defineName',
      'evaluateFormula',
      'format',
      'goTo',
      'goToSpecial',
      'hyperlink',
      'iterative',
      'namedRange',
      'pageSetup',
      'pivotTable',
      'preferences',
    ],
    copied: ['pasteSpecial', 'insertCopiedCells'],
    filled: [],
    misspelt: ['spelling'],
    chart: [],
    pivot: ['pivotFieldSettings', 'pivotOptions'],
  },
};

/** Floating panels over the grid: Escape closes them, with focus as for modal dialogs. */
const FLOATING: Record<Platform, FixtureLists> = {
  default: {
    blank: ['comment', 'filterDropdown', 'findReplace', 'workbookObjects'],
    copied: [],
    filled: ['quickAnalysis'],
    misspelt: [],
    chart: ['chartSettings'],
    pivot: [],
  },
  mac: {
    blank: ['comment', 'filterDropdown', 'findReplace', 'workbookObjects'],
    copied: [],
    filled: [],
    misspelt: [],
    chart: ['chartSettings'],
    pivot: [],
  },
};

/** Docked panes: Escape leaves them open and moves focus to the sheet; only their Close button
 *  closes them. */
const PANES: Record<Platform, FixtureLists> = {
  default: {
    blank: ['watchWindow'],
    copied: [],
    filled: [],
    misspelt: [],
    chart: [],
    pivot: ['pivotFieldList', 'activePivotFieldList'],
  },
  mac: {
    blank: ['functionArguments', 'accessibilityChecker', 'watchWindow', 'externalLinks'],
    copied: [],
    filled: [],
    misspelt: [],
    chart: [],
    pivot: ['pivotFieldList', 'activePivotFieldList'],
  },
};

/** Surfaces a profile does not have: `dialogs.open` throws `unsupported-platform` and `canOpen` is
 *  false. Preferences is the Mac's, and the Mac app has no Quick Analysis. */
const ABSENT: Record<Platform, readonly string[]> = {
  default: ['preferences'],
  mac: ['quickAnalysis'],
};

/** Modals with no Cancel or Close button: the cell style gallery closes by a pick or Escape. */
const NO_CANCEL: readonly string[] = ['cellStylesGallery'];

/** Every id `dialogs` knows, so a new one cannot go unclassified. */
const ALL_IDS = [
  'conditional',
  'iterative',
  'externalLinks',
  'cfRules',
  'cellStylesGallery',
  'evaluateFormula',
  'functionArguments',
  'hyperlink',
  'comment',
  'findReplace',
  'pasteSpecial',
  'insertCopiedCells',
  'namedRange',
  'defineName',
  'pageSetup',
  'format',
  'dataValidation',
  'goTo',
  'goToSpecial',
  'filterDropdown',
  'watchWindow',
  'quickAnalysis',
  'workbookObjects',
  'accessibilityChecker',
  'chartSettings',
  'pivotFieldList',
  'activePivotFieldList',
  'pivotFieldSettings',
  'pivotOptions',
  'pivotTable',
  'spelling',
  'autoCorrect',
  'preferences',
] as const;

/** Every id is a modal dialog, a floating panel, a docked pane or absent on a profile, exactly
 *  once. */
export function expectDialogEscapeClassification(platform: Platform): void {
  const listed = [
    ...Object.values(MODALS[platform]).flat(),
    ...Object.values(FLOATING[platform]).flat(),
    ...Object.values(PANES[platform]).flat(),
    ...ABSENT[platform],
  ];
  expect([...listed].sort(), `${platform}: every dialog id is classified once`).toEqual(
    [...ALL_IDS].sort(),
  );
}

const CONTROLS =
  'input:not([type="hidden"]),select,textarea,button,a[href],[tabindex]:not([tabindex="-1"])';

/** Marks every element showing now, and the focused element, before an open. */
const markBeforeOpen = (page: Page) =>
  page.evaluate(() => {
    const shown = (el: HTMLElement) => !el.hidden && el.getClientRects().length > 0;
    for (const el of Array.from(document.body.querySelectorAll<HTMLElement>('*'))) {
      delete el.dataset.fcEscOpener;
      delete el.dataset.fcEscSurface;
      if (shown(el)) el.dataset.fcEscBefore = '';
      else delete el.dataset.fcEscBefore;
    }
    const active = document.activeElement as HTMLElement | null;
    if (active && active !== document.body) active.dataset.fcEscOpener = '';
  });

/** Tags the surface the open added (the outermost newly shown element holding the most
 *  controls) and focuses its first enabled control. */
const focusNewSurface = (page: Page) =>
  page.evaluate((controls) => {
    const shown = (el: HTMLElement) => !el.hidden && el.getClientRects().length > 0;
    const fresh = Array.from(document.body.querySelectorAll<HTMLElement>('*')).filter(
      (el) => el.dataset.fcEscBefore === undefined && shown(el),
    );
    const freshSet = new Set(fresh);
    const surface = fresh
      .filter((el) => !(el.parentElement && freshSet.has(el.parentElement)))
      .map((el) => ({ el, n: el.querySelectorAll(controls).length }))
      .filter((entry) => entry.n > 0)
      .sort((a, b) => b.n - a.n)[0]?.el;
    for (const el of Array.from(document.querySelectorAll<HTMLElement>('[data-fc-esc-surface]'))) {
      delete el.dataset.fcEscSurface;
    }
    if (!surface) return false;
    // A dock or wrapper can outlive what it holds: take the surface's own root, the element with
    // a surface role holding the most controls (the shallowest on a tie).
    const roles =
      '[role="dialog"],[role="alertdialog"],[role="region"],[role="complementary"],[role="menu"]';
    let root = surface;
    let best = -1;
    for (const el of [surface, ...Array.from(surface.querySelectorAll<HTMLElement>(roles))]) {
      if (!el.matches(roles) || !shown(el)) continue;
      const n = el.querySelectorAll(controls).length;
      if (n > best) {
        best = n;
        root = el;
      }
    }
    root.dataset.fcEscSurface = '';
    const control = Array.from(root.querySelectorAll<HTMLElement>(controls)).find(
      (el) => !(el as HTMLInputElement).disabled && shown(el),
    );
    control?.focus();
    return !!control && document.activeElement === control;
  }, CONTROLS);

const surfaceShown = (page: Page) =>
  page.evaluate(() => {
    const el = document.querySelector<HTMLElement>('[data-fc-esc-surface]');
    return !!el && el.isConnected && !el.hidden && el.getClientRects().length > 0;
  });

/** Where focus is: `opener`, `sheet` (inside the grid region), `body`, or a short description
 *  of anything else. */
const focusPlace = (page: Page) =>
  page.evaluate(() => {
    const active = document.activeElement as HTMLElement | null;
    if (!active || active === document.body || active === document.documentElement) return 'body';
    if (active.closest('.fc-host__grid')) return 'sheet';
    if (active.dataset.fcEscOpener !== undefined) return 'opener';
    return `${active.tagName.toLowerCase()}[${active.getAttribute('aria-label') ?? active.className}]`;
  });

const closeDialog = (page: Page, id: string) =>
  page.evaluate((dialog) => {
    (window as unknown as InstanceWindow).__fcInst.dialogs.close(dialog as never);
  }, id);

const seed = (page: Page, rows: (string | number)[][], selection: string) =>
  page.evaluate(
    ({ table, range }) => {
      const inst = (window as unknown as InstanceWindow).__fcInst;
      inst.workbook.batch(() => {
        const sheet = inst.workbook.sheet(0);
        table.forEach((row, r) => {
          row.forEach((value, c) => {
            if (typeof value === 'number') sheet.setNumber(r, c, value);
            else sheet.setText(r, c, value);
          });
        });
      });
      inst.history.clear();
      inst.selection.set(range);
    },
    { table: rows, range: selection },
  );

/** Builds the fixture through the public API and the ribbon, as a user would. */
async function prepare(
  page: Page,
  fixture: DialogEscapeFixture,
  platform: Platform,
): Promise<void> {
  switch (fixture) {
    case 'blank':
    case 'copied':
      return;
    case 'filled':
      await seed(
        page,
        [
          ['Item', 'Qty'],
          ['a', 1],
          ['b', 2],
        ],
        'A1:B3',
      );
      return;
    case 'misspelt':
      await seed(page, [['teh teh']], 'A1');
      return;
    case 'chart': {
      await seed(
        page,
        [
          ['Month', 'Sales', 'Costs'],
          ['Jan', 10, 6],
          ['Feb', 15, 8],
          ['Mar', 8, 3],
        ],
        'A1:C4',
      );
      await selectRibbonTab(page, 'insert');
      if (platform === 'mac') {
        await page
          .locator('[data-ribbon-command="mac.insert.chartColumn"]:visible')
          .first()
          .click();
        await page
          .locator('[data-ribbon-command="mac.insert.chartColumn"][aria-label="Clustered Column"]')
          .first()
          .click();
      } else {
        await page.locator('[data-ribbon-command="chartInsert"]').first().click();
      }
      await expect
        .poll(() =>
          page.evaluate(
            () =>
              (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).listCharts().length,
          ),
        )
        .toBe(1);
      return;
    }
    case 'pivot': {
      await seed(
        page,
        [
          ['Channel', 'Sales'],
          ['Online', 10],
          ['Retail', 20],
        ],
        'A1:B3',
      );
      await selectRibbonTab(page, 'insert');
      await page
        .locator(
          platform === 'mac'
            ? '[data-ribbon-command="mac.insert.pivotTable"]'
            : '[data-ribbon-command="pivotTableInsert"]',
        )
        .first()
        .click();
      const create = page.getByRole('dialog', { name: 'Create PivotTable' });
      await create.getByRole('button', { name: 'OK', exact: true }).click();
      await expect(create).toBeHidden();
      for (const field of ['Channel', 'Sales']) {
        await page.getByRole('checkbox', { name: field, exact: true }).first().check();
      }
      await closeDialog(page, 'pivotFieldList');
      return;
    }
  }
}

/** Puts focus where a user would be when opening `id` under `fixture`. */
async function focusOpener(page: Page, sp: SpreadsheetPage, fixture: DialogEscapeFixture) {
  if (fixture === 'chart') return;
  if (fixture === 'pivot') {
    // A4 holds the pivot's first row item.
    const cell = await gridCellCenter(page, 3, 0);
    await page.mouse.click(cell.x, cell.y);
    return;
  }
  await sp.focusHost();
  if (fixture === 'filled') {
    await page.evaluate(() =>
      (window as unknown as InstanceWindow).__fcInst.selection.set('A1:B3'),
    );
    return;
  }
  if (fixture === 'copied') {
    await sp.typeIntoActiveCell('17');
    await page.keyboard.press('ArrowUp');
    await sp.shortcut('c');
    await page.keyboard.press('ArrowRight');
  }
}

/** Opens a surface as its fixture expects: Chart Settings from the chart's own button. */
async function open(page: Page, id: string): Promise<void> {
  if (id === 'chartSettings') {
    const button = page.getByRole('button', { name: 'Chart Elements', exact: true }).first();
    await button.evaluate((el) => {
      (el as HTMLElement).dataset.fcEscOpener = '';
    });
    await button.click();
    return;
  }
  await page.evaluate((dialog) => {
    (window as unknown as InstanceWindow).__fcInst.dialogs.open(dialog as never);
  }, id);
}

/** Opens `id` from a host-page button outside the mount, activated by keyboard as a user would,
 *  and checks that closing it by `close` hands focus back to that button. */
async function checkOutsideOpener(page: Page, id: string, close: 'Escape' | 'Cancel') {
  await page.evaluate((dialog) => {
    document.getElementById('fc-esc-outside')?.remove();
    const button = document.createElement('button');
    button.id = 'fc-esc-outside';
    button.textContent = 'Open from the page';
    button.style.cssText = 'position:fixed;right:8px;bottom:8px;z-index:2147483647';
    button.addEventListener('click', () =>
      (window as unknown as InstanceWindow).__fcInst.dialogs.open(dialog as never),
    );
    document.body.appendChild(button);
  }, id);
  const outside = page.locator('#fc-esc-outside');
  await outside.focus();
  await markBeforeOpen(page);
  await page.keyboard.press('Enter');
  let opened = false;
  await expect.soft
    .poll(
      async () => {
        opened = await focusNewSurface(page);
        return opened;
      },
      { message: `${id}: opens from a button outside the spreadsheet` },
    )
    .toBe(true);
  if (opened) {
    if (close === 'Escape') {
      await page.keyboard.press('Escape');
    } else {
      const cancel = page
        .locator('[data-fc-esc-surface]')
        .getByRole('button', { name: /^(Cancel|Close)$/ })
        .last();
      const hasCancel = (await cancel.count()) > 0;
      expect.soft(hasCancel, `${id}: has a Cancel or Close button`).toBe(true);
      if (hasCancel) await cancel.click();
      else await page.keyboard.press('Escape');
    }
    await expect.soft
      .poll(() => surfaceShown(page), { message: `${id}: ${close} closes it` })
      .toBe(false);
    await expect
      .soft(outside, `${id}: ${close} returns focus to the button outside the spreadsheet`)
      .toBeFocused();
  }
  await outside.evaluate((el) => el.remove());
}

/** A ribbon control and the page body are not openers: a modal opened from either hands focus
 *  to the sheet when it closes, by Escape or by Cancel, and so does a floating panel opened from
 *  the ribbon when Escape closes it. */
async function checkNonOpeners(page: Page, sp: SpreadsheetPage, platform: Platform) {
  const command = platform === 'mac' ? 'mac.data.validation' : 'dataValidation';
  for (const from of ['ribbon', 'body'] as const) {
    for (const close of ['Escape', 'Cancel'] as const) {
      await sp.focusHost();
      if (from === 'ribbon') {
        await selectRibbonTab(page, 'data');
        const button = page.locator(`[data-ribbon-command="${command}"]:visible`).first();
        await button.focus();
        await markBeforeOpen(page);
        await button.click();
        // On Mac the button drops a menu; its first item opens the dialog.
        if (platform === 'mac') {
          await page.locator('[data-ribbon-command="mac.data.validation.settings"]').click();
        }
      } else {
        await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
        await markBeforeOpen(page);
        await page.evaluate(() =>
          (window as unknown as InstanceWindow).__fcInst.dialogs.open('dataValidation' as never),
        );
      }
      const label = `Data Validation opened from the ${from}, closed by ${close}`;
      let opened = false;
      await expect.soft
        .poll(
          async () => {
            opened = await focusNewSurface(page);
            return opened;
          },
          { message: `${label}: opens` },
        )
        .toBe(true);
      if (!opened) continue;
      const cancel = page
        .locator('[data-fc-esc-surface]')
        .getByRole('button', { name: /^(Cancel|Close)$/ })
        .last();
      if (close === 'Cancel' && (await cancel.count()) > 0) await cancel.click();
      else await page.keyboard.press('Escape');
      await expect.soft.poll(() => surfaceShown(page), { message: `${label}: closes` }).toBe(false);
      const place = await focusPlace(page);
      expect.soft(place, `${label}: focus goes to the sheet, not ${place}`).toBe('sheet');
      if (await surfaceShown(page)) await closeDialog(page, 'dataValidation');
    }
  }
  const comment = platform === 'mac' ? 'mac.review.newComment' : 'newCommentReview';
  const panels = [
    {
      id: 'findReplace',
      label: 'Find opened from the Find & Select menu',
      tab: 'home',
      open: async () => {
        await page.locator('[data-ribbon-command="findHome"]:visible').first().click();
        await page.getByRole('menuitem', { name: 'Find...', exact: true }).click();
      },
    },
    {
      id: 'comment',
      label: 'the comment editor opened from New Comment',
      tab: 'review',
      open: async () => {
        await page.locator(`[data-ribbon-command="${comment}"]:visible`).first().click();
      },
    },
  ];
  for (const panel of panels) {
    await sp.focusHost();
    // Switch tabs first, so the tab's panel is not taken for the surface the open adds.
    await selectRibbonTab(page, panel.tab);
    await markBeforeOpen(page);
    await panel.open();
    let opened = false;
    await expect.soft
      .poll(
        async () => {
          opened = await focusNewSurface(page);
          return opened;
        },
        { message: `${panel.label}: opens` },
      )
      .toBe(true);
    if (!opened) continue;
    await page.keyboard.press('Escape');
    await expect.soft
      .poll(() => surfaceShown(page), { message: `${panel.label}: Escape closes it` })
      .toBe(false);
    const place = await focusPlace(page);
    expect
      .soft(place, `${panel.label}: Escape moves focus to the sheet, not ${place}`)
      .toBe('sheet');
    if (await surfaceShown(page)) await closeDialog(page, panel.id);
  }
}

/** Escape from the first control of each dialog and pane a fixture makes reachable on a
 *  profile, and each modal opened from a button outside the spreadsheet closed by Escape and by
 *  Cancel, and a modal opened from the ribbon or the page body. Checks are soft, so one run reports every surface that breaks the rule. */
export async function runDialogEscapeScenario(
  page: Page,
  platform: Platform,
  fixture: DialogEscapeFixture,
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ platform });
  await sp.expectNoStub();
  await prepare(page, fixture, platform);

  const check = async (id: string, kind: 'dialog' | 'pane') => {
    await focusOpener(page, sp, fixture);
    await markBeforeOpen(page);
    await open(page, id);
    let opened = false;
    await expect.soft
      .poll(
        async () => {
          opened = await focusNewSurface(page);
          return opened;
        },
        { message: `${id}: opens with a control to focus` },
      )
      .toBe(true);
    if (!opened) return;
    const name = await page.evaluate(() => {
      const el = document.querySelector<HTMLElement>('[data-fc-esc-surface]');
      return el?.getAttribute('aria-label') ?? String(el?.className).split(' ')[0];
    });
    const label = `${id} (${name})`;
    await page.keyboard.press('Escape');
    if (kind === 'dialog') {
      await expect.soft
        .poll(() => surfaceShown(page), { message: `${label}: Escape closes it` })
        .toBe(false);
      const place = await focusPlace(page);
      expect
        .soft(
          ['opener', 'sheet'],
          `${label}: focus returns to its opener or the sheet, not ${place}`,
        )
        .toContain(place);
    } else {
      expect
        .soft(await surfaceShown(page), `${label}: Escape leaves the docked pane open`)
        .toBe(true);
      const place = await focusPlace(page);
      expect.soft(place, `${label}: Escape moves focus to the sheet`).toBe('sheet');
    }
    if (await surfaceShown(page)) await closeDialog(page, id);
    if (await surfaceShown(page)) {
      // Not every pane closes through `dialogs.close`; start the next id from a fresh mount.
      await sp.mount({ platform });
      await prepare(page, fixture, platform);
    }
  };

  for (const id of MODALS[platform][fixture]) await check(id, 'dialog');
  for (const id of MODALS[platform][fixture]) {
    for (const close of ['Escape', 'Cancel'] as const) {
      if (close === 'Cancel' && NO_CANCEL.includes(id)) continue;
      await focusOpener(page, sp, fixture);
      await checkOutsideOpener(page, id, close);
      if (await surfaceShown(page)) await closeDialog(page, id);
      if (await surfaceShown(page)) {
        await sp.mount({ platform });
        await prepare(page, fixture, platform);
      }
    }
  }
  for (const id of FLOATING[platform][fixture]) await check(id, 'dialog');
  for (const id of PANES[platform][fixture]) await check(id, 'pane');

  if (fixture === 'blank') {
    await checkNonOpeners(page, sp, platform);
    for (const id of ABSENT[platform]) {
      const result = await page.evaluate((dialog) => {
        const inst = (window as unknown as InstanceWindow).__fcInst;
        const canOpen = inst.dialogs.canOpen(dialog as never);
        try {
          inst.dialogs.open(dialog as never);
          return { canOpen, failure: null };
        } catch (error) {
          const context = String((error as { context?: unknown }).context ?? '');
          const first = context.split(' ')[0] ?? '';
          return {
            canOpen,
            failure: first.startsWith('failure=') ? first.slice('failure='.length) : String(error),
          };
        }
      }, id);
      expect
        .soft(result, `${id}: not on the ${platform} profile`)
        .toEqual({ canOpen: false, failure: 'unsupported-platform' });
    }
  }
  expect(consoleErrors.read()).toEqual([]);
}
