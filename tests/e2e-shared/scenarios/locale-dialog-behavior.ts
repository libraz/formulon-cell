import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const PROFILES = [
  {
    id: 'en',
    profileToken: 'en_US',
    decimal: '1.5',
    dollar: '$1,234.57',
  },
  {
    id: 'ja',
    profileToken: 'ja_JP',
    decimal: '1.5',
    dollar: '¥1,235',
  },
  {
    id: 'de',
    profileToken: 'de_DE',
    decimal: '1,5',
    dollar: '1.234,57 €',
  },
  {
    id: 'fr',
    profileToken: 'fr_FR',
    decimal: '1,5',
    dollar: '1 234,57 €',
  },
  {
    id: 'zh',
    profileToken: 'zh_CN',
    decimal: '1.5',
    dollar: '¥1,234.57',
  },
  {
    id: 'ko',
    profileToken: 'ko_KR',
    decimal: '1.5',
    dollar: '₩1,235',
  },
  {
    id: 'th',
    profileToken: 'th_TH',
    decimal: '1.5',
    dollar: '฿1,234.57',
  },
] as const;

type Platform = 'mac' | 'default';

type ConditionalRuleProbe = {
  kind: string;
  a?: number | string;
  formula?: string;
};

type SpreadsheetGlobal = {
  __fcInst?: {
    locale?: string;
    i18n?: { locale?: string };
    workbook?: {
      spreadsheetProfileId?: () => string | null;
      getValue: (addr: { sheet: number; row: number; col: number }) => unknown;
      save: () => Uint8Array;
    };
    store?: {
      getState: () => {
        conditional: { rules: ConditionalRuleProbe[] };
      };
    };
    openConditionalDialog?: (options?: {
      mode?: 'manage' | 'new' | 'edit';
      editIndex?: number;
      kind?: 'cell-value' | 'formula';
      cellValueOp?: '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between';
    }) => void;
    openWatchWindow?: () => void;
    setFeatures?: (flags: { localeSwitcher?: boolean }) => void;
  };
};

const readInstanceState = (page: Page) =>
  page.evaluate(() => {
    const instance = (window as unknown as SpreadsheetGlobal).__fcInst;
    if (!instance?.workbook) throw new Error('Missing spreadsheet instance');
    return {
      locale: instance.locale ?? instance.i18n?.locale ?? null,
      profile: instance.workbook.spreadsheetProfileId?.() ?? null,
    };
  });

const readCell = (page: Page, row: number, col = 0) =>
  page.evaluate(
    ({ row: targetRow, col: targetCol }) => {
      const instance = (window as unknown as SpreadsheetGlobal).__fcInst;
      if (!instance?.workbook) throw new Error('Missing spreadsheet instance');
      return instance.workbook.getValue({ sheet: 0, row: targetRow, col: targetCol });
    },
    { row, col },
  );

const readRules = (page: Page): Promise<ConditionalRuleProbe[]> =>
  page.evaluate(() => {
    const instance = (window as unknown as SpreadsheetGlobal).__fcInst;
    const rules = instance?.store?.getState().conditional.rules ?? [];
    return rules.map((rule) => ({
      kind: rule.kind,
      ...(rule.a !== undefined ? { a: rule.a } : {}),
      ...(rule.formula !== undefined ? { formula: rule.formula } : {}),
    }));
  });

const goTo = async (page: Page, ref: string): Promise<void> => {
  const nameBox = page.locator('.fc-host__formulabar-tag').first();
  await nameBox.fill(ref);
  await nameBox.press('Enter');
  await expect(page.locator('.fc-host').first()).toBeFocused();
  await expect.poll(() => nameBox.inputValue()).toBe(ref);
};

const enableLocaleSwitcher = async (page: Page) => {
  await page.evaluate(() => {
    (window as unknown as SpreadsheetGlobal).__fcInst?.setFeatures?.({ localeSwitcher: true });
  });
  const selector = page.locator('.fc-host__locale-switcher-select').first();
  await expect(selector).toBeVisible();
  await expect(selector.locator('option')).toHaveCount(PROFILES.length);
  return selector;
};

const switchLocale = async (
  page: Page,
  selector: ReturnType<Page['locator']>,
  profile: (typeof PROFILES)[number],
  platform: Platform,
): Promise<void> => {
  await selector.selectOption(profile.id);
  await expect(selector).toHaveValue(profile.id);
  await expect(page.locator('html')).toHaveAttribute('lang', profile.id);
  await expect(page.locator('.fc-host').first()).toHaveAttribute('lang', profile.id);
  await expect(page.locator('.fc-overlay-portal').first()).toHaveAttribute('lang', profile.id);
  await expect
    .poll(() => readInstanceState(page), { message: `active locale should be ${profile.id}` })
    .toEqual({
      locale: profile.id,
      profile: `${platform === 'mac' ? 'mac' : 'windows'}-${profile.profileToken}`,
    });
};

const openConditional = async (
  page: Page,
  options: Parameters<NonNullable<SpreadsheetGlobal['__fcInst']>['openConditionalDialog']>[0],
) => {
  await page.evaluate((dialogOptions) => {
    const instance = (window as unknown as SpreadsheetGlobal).__fcInst;
    if (!instance?.openConditionalDialog) throw new Error('Conditional dialog API unavailable');
    instance.openConditionalDialog(dialogOptions);
  }, options);
  const dialog = page.locator('.fc-conddlg').first();
  await expect(dialog).toBeVisible();
  return dialog;
};

const conditionalValueInput = (dialog: ReturnType<Page['locator']>) =>
  dialog.locator('input[type="text"]').nth(1);

const conditionalFormulaInput = (dialog: ReturnType<Page['locator']>) =>
  dialog.locator('input[placeholder*="="]').first();

const addConditionalRule = async (dialog: ReturnType<Page['locator']>): Promise<void> => {
  await dialog.locator('.fc-conddlg__addrow button').click();
  await expect(dialog).toBeHidden();
};

type LocaleScenarioContext = {
  sp: UserJourneyPage;
  localeSelect: ReturnType<Page['locator']>;
  consoleErrors: { read: () => string[] };
};

const setupLocaleScenario = async (
  page: Page,
  platform: Platform,
): Promise<LocaleScenarioContext> => {
  const sp = new UserJourneyPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ locale: 'en', fixture: 'empty', platform });
  await sp.expectNoStub();
  const localeSelect = await enableLocaleSwitcher(page);
  return { sp, localeSelect, consoleErrors };
};

const finishLocaleScenario = async (
  page: Page,
  consoleErrors: LocaleScenarioContext['consoleErrors'],
): Promise<void> => {
  await page.evaluate(() => {
    (window as unknown as SpreadsheetGlobal).__fcInst?.setFeatures?.({ localeSwitcher: false });
  });
  await expect(page.locator('.fc-host__locale-switcher-select')).toHaveCount(0);
  expect(consoleErrors.read(), 'locale dialog scenario should not emit console errors').toEqual([]);
};

/**
 * I07 — verify conditional rule drafts and numeric criteria across locales.
 *
 * Keep this flow independent so it gets a fresh workbook and its seven-rule
 * hydration loop stays within Playwright's normal per-test deadline.
 */
export async function runConditionalLocaleDialogScenario(
  page: Page,
  platform: Platform = 'default',
): Promise<void> {
  const { sp, localeSelect, consoleErrors } = await setupLocaleScenario(page, platform);

  // A real editable dialog keeps its draft while locale changes are rejected.
  // The selector is force-clicked because the modal backdrop correctly owns
  // the hit target; the event still goes through the switcher's DOM handler.
  await sp.focusHost();
  const modal = await openConditional(page, {
    mode: 'new',
    kind: 'cell-value',
    cellValueOp: '>',
  });
  const modalValue = conditionalValueInput(modal);
  await modalValue.fill('1.5');
  await localeSelect.selectOption('de', { force: true });
  await expect(localeSelect).toHaveValue('en');
  await expect
    .poll(() => readInstanceState(page))
    .toEqual({
      locale: 'en',
      profile: `${platform === 'mac' ? 'mac' : 'windows'}-en_US`,
    });
  await expect(page.locator('.fc-host').first()).toHaveAttribute('lang', 'en');
  await expect(page.locator('.fc-overlay-portal').first()).toHaveAttribute('lang', 'en');
  await expect(modalValue).toHaveValue('1.5');
  await page.keyboard.press('Escape');
  await expect(modal).toBeHidden();

  // The same active editor can now be switched after it is closed.
  await switchLocale(page, localeSelect, PROFILES[2], platform);

  // Every profile must parse its own decimal separator, persist a canonical
  // numeric rule, and hydrate that rule back into the active locale.
  for (const profile of PROFILES) {
    await switchLocale(page, localeSelect, profile, platform);
    await goTo(page, 'A1');
    await page.keyboard.type(profile.decimal);
    await page.keyboard.press('Enter');
    await page.keyboard.press('ArrowUp');
    await expect.poll(() => readCell(page, 0)).toEqual({ kind: 'number', value: 1.5 });
    await expect.poll(() => sp.formulaBarValue()).toBe(profile.decimal);

    const dialog = await openConditional(page, {
      mode: 'new',
      kind: 'cell-value',
      cellValueOp: '>',
    });
    await conditionalValueInput(dialog).fill(profile.decimal);
    await addConditionalRule(dialog);
    await expect.poll(() => readRules(page)).toContainEqual({ kind: 'cell-value', a: 1.5 });

    const rules = await readRules(page);
    const editIndex = rules.length - 1;
    const editDialog = await openConditional(page, { mode: 'edit', editIndex });
    await expect(conditionalValueInput(editDialog)).toHaveValue(profile.decimal);
    await page.keyboard.press('Escape');
    await expect(editDialog).toBeHidden();

    // Keep every rule in the workbook through the rest of the locale loop.
    // This makes a later switch exercise the renderer's profile invalidation
    // with unchanged cell values and already-authored conditional state.
  }

  await finishLocaleScenario(page, consoleErrors);
}

/** I07 — verify localized conditional formulas and the workbook save path. */
export async function runConditionalFormulaLocaleScenario(
  page: Page,
  platform: Platform = 'default',
): Promise<void> {
  const { sp, localeSelect, consoleErrors } = await setupLocaleScenario(page, platform);

  // Localized function names and separators are converted only at the formula
  // input boundary. The stored rule remains the canonical AND expression,
  // which is also what the xlsx save path consumes.
  await sp.enter('A1', '1.5');
  await expect.poll(() => readCell(page, 0)).toEqual({ kind: 'number', value: 1.5 });
  await switchLocale(page, localeSelect, PROFILES[2], platform);
  await goTo(page, 'A1');
  const germanFormulaDialog = await openConditional(page, {
    mode: 'new',
    kind: 'formula',
  });
  const kindSelect = germanFormulaDialog.locator('select:has(option[value="formula"])');
  await kindSelect.selectOption('formula');
  await conditionalFormulaInput(germanFormulaDialog).fill('=UND(A1>1,5;A1<3;A1<>"a;b")');
  await addConditionalRule(germanFormulaDialog);
  let rules = await readRules(page);
  const germanFormulaIndex = rules.length - 1;
  expect(rules[germanFormulaIndex]).toEqual({
    kind: 'formula',
    formula: '=AND(A1>1.5,A1<3,A1<>"a;b")',
  });

  const germanEdit = await openConditional(page, {
    mode: 'edit',
    editIndex: germanFormulaIndex,
  });
  await expect(conditionalFormulaInput(germanEdit)).toHaveValue('=UND(A1>1,5;A1<3;A1<>"a;b")');
  await page.keyboard.press('Escape');
  await expect(germanEdit).toBeHidden();

  await switchLocale(page, localeSelect, PROFILES[3], platform);
  const frenchEdit = await openConditional(page, {
    mode: 'edit',
    editIndex: germanFormulaIndex,
  });
  await expect(conditionalFormulaInput(frenchEdit)).toHaveValue('=ET(A1>1,5;A1<3;A1<>"a;b")');
  await page.keyboard.press('Escape');
  await expect(frenchEdit).toBeHidden();

  await switchLocale(page, localeSelect, PROFILES[0], platform);
  const englishEdit = await openConditional(page, {
    mode: 'edit',
    editIndex: germanFormulaIndex,
  });
  await expect(conditionalFormulaInput(englishEdit)).toHaveValue('=AND(A1>1.5,A1<3,A1<>"a;b")');
  await page.keyboard.press('Escape');
  await expect(englishEdit).toBeHidden();

  const saved = await page.evaluate(() => {
    const instance = (window as unknown as SpreadsheetGlobal).__fcInst;
    if (!instance?.workbook) throw new Error('Missing workbook save API');
    const bytes = instance.workbook.save();
    return { length: bytes.byteLength, zip: bytes[0] === 0x50 && bytes[1] === 0x4b };
  });
  expect(saved.zip).toBe(true);
  expect(saved.length).toBeGreaterThan(100);
  rules = await readRules(page);
  expect(rules[germanFormulaIndex]).toEqual({
    kind: 'formula',
    formula: '=AND(A1>1.5,A1<3,A1<>"a;b")',
  });

  await finishLocaleScenario(page, consoleErrors);
}

/** I07 — exercise Mac Goal Seek input and empty-target validation in German. */
export async function runMacGoalSeekLocaleDialogScenario(page: Page): Promise<void> {
  const platform: Platform = 'mac';
  const { sp, localeSelect, consoleErrors } = await setupLocaleScenario(page, platform);

  // Mac Goal Seek accepts a localized decimal through a text input. Exercise
  // the physical keyboard path once on the German profile, then verify that
  // an empty draft is rejected without being coerced to zero.
  await switchLocale(page, localeSelect, PROFILES[2], platform);
  await sp.enter('E1', '3');
  await sp.enter('F1', '=E1*2');
  await sp.goTo('F1');
  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('mac.data.whatIf');
  await page.locator('#menu-mac-data-what-if [data-ribbon-command="mac.data.goalSeek"]').click();
  const goalSeek = page.locator('.fc-mac-goalseek');
  await expect(goalSeek).toBeVisible();
  await goalSeek.locator('#fc-mac-goalseek-formula-cell').fill('F1');
  await goalSeek.locator('#fc-mac-goalseek-changing-cell').fill('E1');
  const target = goalSeek.locator('#fc-mac-goalseek-target-value');
  await target.fill('');
  await target.pressSequentially('1,5');
  await expect(target).toHaveValue('1,5');
  const goalSeekSubmit = goalSeek.locator('[data-fc-mac-action="goal-seek-ok"]');
  await goalSeekSubmit.click();
  await expect(goalSeek.locator('.fc-mac-goalseek__status')).toHaveAttribute(
    'data-state',
    'success',
  );
  await goalSeekSubmit.click();
  await expect(goalSeek).toBeHidden();
  await expect.poll(() => sp.readValue('E1')).toEqual({ kind: 'number', value: 0.75 });
  await expect.poll(() => sp.readValue('F1')).toEqual({ kind: 'number', value: 1.5 });

  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('mac.data.whatIf');
  await page.locator('#menu-mac-data-what-if [data-ribbon-command="mac.data.goalSeek"]').click();
  await expect(goalSeek).toBeVisible();
  const emptyTarget = goalSeek.locator('#fc-mac-goalseek-target-value');
  await emptyTarget.fill('');
  await goalSeek.locator('[data-fc-mac-action="goal-seek-ok"]').click();
  await expect(goalSeek.locator('.fc-mac-goalseek__status')).toHaveAttribute('data-state', 'error');
  await expect(emptyTarget).toHaveValue('');
  await goalSeek.locator('[data-fc-mac-action="goal-seek-cancel"]').click();
  await expect(goalSeek).toBeHidden();

  await finishLocaleScenario(page, consoleErrors);
}

/** I07 — verify Watch Window values stay localized while the window remains open. */
export async function runWatchWindowLocaleDialogScenario(
  page: Page,
  platform: Platform = 'default',
): Promise<void> {
  const { sp, localeSelect, consoleErrors } = await setupLocaleScenario(page, platform);

  // Watch Window uses the profile-aware display renderer and stays mounted
  // while the active locale changes. Add the row through its visible button so
  // the test covers the same path as the Review ribbon command.
  await sp.goTo('C1');
  await page.keyboard.type('=DOLLAR(1234.567)');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.evaluate(() => {
    const instance = (window as unknown as SpreadsheetGlobal).__fcInst;
    if (!instance?.openWatchWindow) throw new Error('Watch Window API unavailable');
    instance.openWatchWindow();
  });
  const watch = page.locator('.fc-watch').first();
  await expect(watch).toBeVisible();
  await watch.locator('.fc-watch__btn').first().click();
  const watchRow = watch.locator('.fc-watch__row');
  await expect(watchRow).toHaveCount(1);
  await expect(watchRow.locator('td').nth(3)).toHaveText('C1');
  await expect(watchRow.locator('td').nth(0)).not.toBeEmpty();
  await expect(watchRow.locator('td').nth(1)).not.toBeEmpty();

  for (const profile of PROFILES) {
    await switchLocale(page, localeSelect, profile, platform);
    await expect(watchRow).toHaveCount(1);
    await expect(watchRow.locator('.fc-watch__value')).toHaveText(profile.dollar);
  }
  await watch.locator('.fc-watch__close').click();
  await expect(watch).toBeHidden();

  await finishLocaleScenario(page, consoleErrors);
}
