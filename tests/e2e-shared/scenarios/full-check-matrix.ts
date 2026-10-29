import { expect, type Page } from '@playwright/test';

import matrix from '../fixtures/full-check-matrix.json' with { type: 'json' };
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export const FULL_CHECK_LOCALES = ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const;
export type FullCheckLocale = (typeof FULL_CHECK_LOCALES)[number];
export type FullCheckPlatform = 'default' | 'mac';
export type FullCheckTheme = 'paper' | 'ink' | 'contrast';
export type FullCheckFlow = 'objects' | 'clipboard' | 'filter' | 'formula' | 'analysis' | 'menus';

export interface FullCheckRow {
  locale: FullCheckLocale;
  platform: FullCheckPlatform;
  theme: FullCheckTheme;
  width: '640' | '1440';
  flow: FullCheckFlow;
}

export const FULL_CHECK_MATRIX = matrix.tests as FullCheckRow[];

const PROFILE_TOKENS: Record<FullCheckLocale, string> = {
  en: 'en_US',
  ja: 'ja_JP',
  de: 'de_DE',
  fr: 'fr_FR',
  zh: 'zh_CN',
  ko: 'ko_KR',
  th: 'th_TH',
};

const DEMO_LOCALE_LABELS: Record<FullCheckLocale, string> = {
  en: 'English',
  ja: '日本語',
  de: 'Deutsch',
  fr: 'Français',
  zh: '中文',
  ko: '한국어',
  th: 'ไทย',
};

const LOCALE_TAGS: Record<FullCheckLocale, string> = {
  en: 'en-US',
  ja: 'ja-JP',
  de: 'de-DE',
  fr: 'fr-FR',
  zh: 'zh-CN',
  ko: 'ko-KR',
  th: 'th-TH',
};

const localizedDecimal = (locale: FullCheckLocale, value: number): string =>
  new Intl.NumberFormat(LOCALE_TAGS[locale]).format(value);

const localizedBoolean = (locale: FullCheckLocale, value: boolean): string => {
  if (locale === 'de') return value ? 'WAHR' : 'FALSCH';
  if (locale === 'fr') return value ? 'VRAI' : 'FAUX';
  return value ? 'TRUE' : 'FALSE';
};

const localeAfter = (locale: FullCheckLocale): FullCheckLocale => {
  const index = FULL_CHECK_LOCALES.indexOf(locale);
  return FULL_CHECK_LOCALES[(index + 1) % FULL_CHECK_LOCALES.length] ?? 'en';
};

type CellAddress = { sheet: number; row: number; col: number };
type SeedCell = {
  address: CellAddress;
  kind: 'text' | 'number' | 'bool' | 'formula';
  value: string | number | boolean;
};

type BrowserInstance = {
  locale: string;
  i18n: { locale: string; strings: Record<string, unknown> };
  workbook: {
    capabilities: Record<string, boolean>;
    setText(address: CellAddress, value: string): void;
    setNumber(address: CellAddress, value: number): void;
    setBool(address: CellAddress, value: boolean): void;
    setFormula(address: CellAddress, value: string): void;
    recalc(): void;
    getValue(address: CellAddress): unknown;
    cellFormula(address: CellAddress): string | null;
    getEffectiveStyle?(
      sheet: number,
      row: number,
      col: number,
    ): { xfIndex: number; source?: number } | null;
    getCellXf?(xfIndex: number): { fontIndex: number; [key: string]: unknown } | null;
    getFontRecord?(
      fontIndex: number,
    ): { name: string; size: number; [key: string]: unknown } | null;
    setWorkbookDefaultFont?(record: { name?: string; size?: number }): boolean;
    addFontRecord?(record: Record<string, unknown>): number;
    addXfRecord?(record: Record<string, unknown>): number;
    setColumnStyle?(sheet: number, first: number, last: number, xfIndex: number): boolean;
    setRowStyle?(sheet: number, row: number, xfIndex: number): boolean;
    spreadsheetProfileId?(): string | null;
    getTables?(): { name: string; displayName: string; ref: string; columns: string[] }[];
    setSheetAutoFilterXml?(sheet: number, xml: string): boolean;
  };
  history: unknown;
  toolbar?: {
    getDisplayMode?(): string;
  } | null;
  store: {
    getState(): {
      selection: unknown;
      ui: {
        filterRanges: unknown[];
        filterCriteria: unknown[];
        editor: { kind: string };
        copyMode?: string | null;
      };
      layout: { hiddenRows: Set<number> };
      charts: { charts: unknown[] };
      conditional: { rules: unknown[] };
      errorIndicators: { ignoredErrors: Set<string> };
    };
    setState(updater: (state: Record<string, unknown>) => Record<string, unknown>): void;
  };
  setLocale(locale: string): boolean;
  setFeatures(next: Record<string, boolean>): void;
  setTheme(theme: FullCheckTheme): void;
  clipboard?: {
    getSnapshot(): {
      cells: Array<Array<{ format?: { fontFamily?: string; fontSize?: number } }>>;
    } | null;
  } | null;
  openPivotTableDialog(): void;
  openActivePivotFieldList(): boolean;
  addSlicer(input: {
    tableName: string;
    column: string;
    selected?: readonly string[];
    x?: number;
    y?: number;
  }): { id: string };
  openPasteSpecial(): void;
  openInsertCopiedCells(): void;
  openFilterDropdown(
    range: { sheet: number; r0: number; c0: number; r1: number; c1: number },
    col: number,
  ): void;
  openQuickAnalysis(): void;
};

type ProbeWindow = Window & {
  __fcInst?: BrowserInstance;
  __fcToolbar?: { getDisplayMode?(): string } | null;
};

type RuntimeSnapshot = {
  locale: string;
  profile: string | null;
  filterRanges: unknown[];
  filterCriteria: unknown[];
  hiddenRows: number[];
  conditionalRuleCount: number;
  ignoredErrorCount: number;
  chartCount: number;
};

type NativeFont = { name: string; size: number; source?: number };
type ErrorTriangleProbe = {
  addr: { sheet: number; row: number; col: number };
  rect: { x: number; y: number; w: number; h: number };
};

interface SurfaceLabels {
  home: string;
  profile: string | null;
  contextTitle: string;
  copy: string;
  paste: string;
  pasteSpecial: string;
  insertCopiedCells: string;
  pasteSpecialTitle: string;
  pasteCancel: string;
  insertCopiedTitle: string;
  insertCancel: string;
  pasteOptionsTitle: string;
  pasteKeepSource: string;
  pasteValues: string;
  pasteFormatting: string;
  filterTitle: string;
  filterSearch: string;
  filterSelectAll: string;
  filterBlanks: string;
  filterApply: string;
  filterClear: string;
  autocompletePick: string;
  autocompleteCustom: string;
  argHelperIntersection: string;
  quickTitle: string;
  quickDataBar: string;
  errorHeading: string;
  errorShowInfo: string;
  errorIgnore: string;
  pivotFieldList: string;
  slicerTitle: string;
  slicerClear: string;
  slicerClose: string;
  blank: string;
  chartElements: string;
  backstageFile: string;
  backstageInfo: string;
  backstageClose: string;
  displayLabel: string;
  displayExpanded: string;
  displaySingleLine: string;
  displayCollapsed: string;
  displayAutoHide: string;
}

const demoStringsUrl = `/@fs${new URL('../../../apps/demo-shared/demo-strings.ts', import.meta.url).pathname}`;

const getStrings = async (page: Page): Promise<SurfaceLabels> =>
  page.evaluate(async (moduleUrl) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check requires window.__fcInst');
    const { createDemoStrings } = await import(moduleUrl);
    const demoStrings =
      createDemoStrings('React')[instance.locale as keyof ReturnType<typeof createDemoStrings>];
    if (!demoStrings) throw new Error(`full-check demo strings missing for ${instance.locale}`);
    const strings = instance.i18n.strings as unknown as {
      ribbon: {
        tabs: { home: string; file: string };
      };
      backstage: { info: string; close: string };
      contextMenu: {
        title: string;
        copy: string;
        paste: string;
        pasteSpecial: string;
        insertCopiedCells: string;
      };
      macContextMenu: { pasteSpecial: string };
      pasteSpecialDialog: { title: string; cancel: string };
      insertCopiedCellsDialog: { title: string; cancel: string };
      pasteOptions: {
        title: string;
        keepSourceFormatting: string;
        values: string;
        formattingOnly: string;
      };
      filterDropdown: {
        title: string;
        searchPlaceholder: string;
        selectAll: string;
        blanks: string;
        apply: string;
        clear: string;
      };
      autocomplete: { pickFromList: string; customFunction: string };
      argHelper: { implicitIntersection: string };
      quickAnalysis: { title: string; actions: { dataBar: string } };
      errorMenu: { errorHeading: string; showInfo: string; ignore: string };
      pivotTableDialog: { fieldList: string };
      slicer: { title: string; clear: string; close: string };
      workbookObjects: { pivotBlankItem: string };
      sessionCharts: { columnChart: string };
      ribbonDisplay: {
        label: string;
        expanded: string;
        singleLine: string;
        collapsed: string;
        autoHide: string;
      };
    };
    return {
      home: strings.ribbon.tabs.home,
      profile: instance.workbook.spreadsheetProfileId?.() ?? null,
      contextTitle: strings.contextMenu.title,
      copy: strings.contextMenu.copy,
      paste: strings.contextMenu.paste,
      // The Mac menu opens Paste Special as a submenu, so its label has no ellipsis.
      pasteSpecial:
        document.querySelector<HTMLElement>('.fc-host')?.dataset.fcPlatform === 'mac'
          ? strings.macContextMenu.pasteSpecial
          : strings.contextMenu.pasteSpecial,
      insertCopiedCells: strings.contextMenu.insertCopiedCells,
      pasteSpecialTitle: strings.pasteSpecialDialog.title,
      pasteCancel: strings.pasteSpecialDialog.cancel,
      insertCopiedTitle: strings.insertCopiedCellsDialog.title,
      insertCancel: strings.insertCopiedCellsDialog.cancel,
      pasteOptionsTitle: strings.pasteOptions.title,
      pasteKeepSource: strings.pasteOptions.keepSourceFormatting,
      pasteValues: strings.pasteOptions.values,
      pasteFormatting: strings.pasteOptions.formattingOnly,
      filterTitle: strings.filterDropdown.title,
      filterSearch: strings.filterDropdown.searchPlaceholder,
      filterSelectAll: strings.filterDropdown.selectAll,
      filterBlanks: strings.filterDropdown.blanks,
      filterApply: strings.filterDropdown.apply,
      filterClear: strings.filterDropdown.clear,
      autocompletePick: strings.autocomplete.pickFromList,
      autocompleteCustom: strings.autocomplete.customFunction,
      argHelperIntersection: strings.argHelper.implicitIntersection,
      quickTitle: strings.quickAnalysis.title,
      quickDataBar: strings.quickAnalysis.actions.dataBar,
      errorHeading: strings.errorMenu.errorHeading,
      errorShowInfo: strings.errorMenu.showInfo,
      errorIgnore: strings.errorMenu.ignore,
      pivotFieldList: strings.pivotTableDialog.fieldList,
      slicerTitle: strings.slicer.title,
      slicerClear: strings.slicer.clear,
      slicerClose: strings.slicer.close,
      blank: strings.workbookObjects.pivotBlankItem,
      chartElements: strings.sessionCharts.elementsButton,
      backstageFile: demoStrings.file,
      backstageInfo: demoStrings.info,
      backstageClose: demoStrings.close,
      displayLabel: strings.ribbonDisplay.label,
      displayExpanded: strings.ribbonDisplay.expanded,
      displaySingleLine: strings.ribbonDisplay.singleLine,
      displayCollapsed: strings.ribbonDisplay.collapsed,
      displayAutoHide: strings.ribbonDisplay.autoHide,
    } satisfies SurfaceLabels;
  }, demoStringsUrl);

const runtime = (page: Page): Promise<RuntimeSnapshot> =>
  page.evaluate(() => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check requires a mounted spreadsheet instance');
    const state = instance.store.getState();
    return {
      locale: instance.locale,
      profile: instance.workbook.spreadsheetProfileId?.() ?? null,
      filterRanges: state.ui.filterRanges,
      filterCriteria: state.ui.filterCriteria,
      hiddenRows: [...state.layout.hiddenRows].sort((a, b) => a - b),
      conditionalRuleCount: state.conditional.rules.length,
      ignoredErrorCount: state.errorIndicators.ignoredErrors.size,
      chartCount: state.charts.charts.length,
    };
  });

const setSelection = async (
  page: Page,
  range: { sheet: number; r0: number; c0: number; r1: number; c1: number },
): Promise<void> => {
  await page.evaluate((nextRange) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check selection requires __fcInst');
    instance.store.setState((state) => ({
      ...state,
      selection: {
        active: { sheet: nextRange.sheet, row: nextRange.r0, col: nextRange.c0 },
        anchor: { sheet: nextRange.sheet, row: nextRange.r0, col: nextRange.c0 },
        range: nextRange,
        extraRanges: [],
      },
    }));
  }, range);
};

const seed = async (
  page: Page,
  cells: readonly SeedCell[],
  selection?: { sheet: number; r0: number; c0: number; r1: number; c1: number },
): Promise<void> => {
  await page.evaluate(
    ({ cells: nextCells, selection: nextSelection }) => {
      const instance = (window as ProbeWindow).__fcInst;
      if (!instance) throw new Error('full-check seed requires __fcInst');
      for (const cell of nextCells) {
        if (cell.kind === 'text') instance.workbook.setText(cell.address, String(cell.value));
        else if (cell.kind === 'number')
          instance.workbook.setNumber(cell.address, Number(cell.value));
        else if (cell.kind === 'bool') instance.workbook.setBool(cell.address, Boolean(cell.value));
        else instance.workbook.setFormula(cell.address, String(cell.value));
      }
      instance.workbook.recalc();
      if (nextSelection) {
        instance.store.setState((state) => ({
          ...state,
          selection: {
            active: { sheet: nextSelection.sheet, row: nextSelection.r0, col: nextSelection.c0 },
            anchor: { sheet: nextSelection.sheet, row: nextSelection.r0, col: nextSelection.c0 },
            range: nextSelection,
            extraRanges: [],
          },
        }));
      }
    },
    { cells, selection },
  );
};

const readCells = async (page: Page, addresses: readonly CellAddress[]): Promise<unknown[]> =>
  page.evaluate((nextAddresses) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check read requires __fcInst');
    return nextAddresses.map((address) => instance.workbook.getValue(address));
  }, addresses);

const readCellFormula = async (page: Page, address: CellAddress): Promise<string | null> =>
  page.evaluate((nextAddress) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check formula read requires __fcInst');
    return instance.workbook.cellFormula(nextAddress);
  }, address);

const readNativeFonts = async (
  page: Page,
  addresses: readonly CellAddress[],
): Promise<NativeFont[]> =>
  page.evaluate((nextAddresses) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check native font read requires __fcInst');
    const workbook = instance.workbook;
    if (
      !workbook.getEffectiveStyle ||
      !workbook.getCellXf ||
      !workbook.getFontRecord ||
      workbook.capabilities.effectiveStyle !== true
    ) {
      throw new Error(
        'full-check clipboard fixture unavailable: engine effectiveStyle capability is false',
      );
    }
    return nextAddresses.map((address) => {
      const effective = workbook.getEffectiveStyle?.(address.sheet, address.row, address.col);
      const xf = effective ? workbook.getCellXf?.(effective.xfIndex) : null;
      const font = xf ? workbook.getFontRecord?.(xf.fontIndex) : null;
      if (!font || typeof font.name !== 'string' || !Number.isFinite(font.size)) {
        throw new Error(
          `full-check native font read failed at ${address.sheet}:${address.row}:${address.col}`,
        );
      }
      return {
        name: font.name,
        size: font.size,
        ...(typeof effective?.source === 'number' ? { source: effective.source } : {}),
      };
    });
  }, addresses);

const readClipboardSnapshotFonts = async (
  page: Page,
  offsets: readonly { row: number; col: number }[],
): Promise<Array<{ fontFamily: string; fontSize: number }>> =>
  page.evaluate((nextOffsets) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check clipboard snapshot requires __fcInst');
    const snapshot = instance.clipboard?.getSnapshot();
    if (!snapshot) throw new Error('full-check clipboard snapshot is unavailable after copy');
    return nextOffsets.map(({ row, col }) => {
      const format = snapshot.cells[row]?.[col]?.format;
      const fontFamily = format?.fontFamily;
      const fontSize = format?.fontSize;
      if (
        typeof fontFamily !== 'string' ||
        typeof fontSize !== 'number' ||
        !Number.isFinite(fontSize)
      ) {
        throw new Error(`full-check clipboard snapshot omitted effective font at ${row}:${col}`);
      }
      return { fontFamily, fontSize };
    });
  }, offsets);

const setNativeInheritedClipboardFonts = async (page: Page): Promise<NativeFont[]> => {
  const sourceAddresses: CellAddress[] = [
    { sheet: 0, row: 0, col: 3 },
    { sheet: 0, row: 0, col: 4 },
    { sheet: 0, row: 1, col: 3 },
  ];
  const fonts = await page.evaluate((nextAddresses) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check native font fixture requires __fcInst');
    const workbook = instance.workbook;
    if (
      workbook.capabilities.defaultFont !== true ||
      workbook.capabilities.sheetFontStyles !== true ||
      workbook.capabilities.cellFormatting !== true ||
      !workbook.setWorkbookDefaultFont ||
      !workbook.getCellXf ||
      !workbook.getFontRecord ||
      !workbook.addFontRecord ||
      !workbook.addXfRecord ||
      !workbook.setColumnStyle ||
      !workbook.setRowStyle
    ) {
      throw new Error(
        'full-check clipboard fixture unavailable: native inherited font capabilities are incomplete',
      );
    }

    const makeAxisXf = (size: number): number => {
      const baseXf = workbook.getCellXf?.(0);
      const baseFont = baseXf ? workbook.getFontRecord?.(baseXf.fontIndex) : null;
      if (!baseXf || !baseFont) {
        throw new Error('full-check clipboard fixture failed: default native font is unavailable');
      }
      const fontIndex = workbook.addFontRecord?.({ ...baseFont, size });
      if (fontIndex === undefined || fontIndex < 0) {
        throw new Error('full-check clipboard fixture failed: row/column font was not created');
      }
      const xfIndex = workbook.addXfRecord?.({ ...baseXf, fontIndex });
      if (xfIndex === undefined || xfIndex < 0) {
        throw new Error('full-check clipboard fixture failed: row/column XF was not created');
      }
      return xfIndex;
    };

    // D1 resolves through the workbook default, E1 through a native column
    // XF, and D2 through a native row XF. These source cells stay blank so
    // they cannot acquire a direct cell XF from value seeding; the clipboard
    // resolver must capture their effective styles rather than silently
    // dropping the inherited font axis.
    if (!workbook.setWorkbookDefaultFont({ size: 15 })) {
      throw new Error('full-check clipboard fixture failed: workbook default font was rejected');
    }
    const columnXf = makeAxisXf(17);
    const rowXf = makeAxisXf(19);
    if (!workbook.setColumnStyle(0, 4, 4, columnXf)) {
      throw new Error('full-check clipboard fixture failed: column font was rejected');
    }
    if (!workbook.setRowStyle(0, 1, rowXf)) {
      throw new Error('full-check clipboard fixture failed: row font was rejected');
    }

    const fonts = nextAddresses.map((address) => {
      const effective = workbook.getEffectiveStyle?.(address.sheet, address.row, address.col);
      const xf = effective ? workbook.getCellXf?.(effective.xfIndex) : null;
      const font = xf ? workbook.getFontRecord?.(xf.fontIndex) : null;
      if (!font || typeof font.name !== 'string' || !Number.isFinite(font.size)) {
        throw new Error(
          `full-check clipboard fixture failed: effective font missing at ${address.row}:${address.col}`,
        );
      }
      if (typeof effective?.source !== 'number') {
        throw new Error(
          `full-check clipboard fixture failed: inherited source missing at ${address.row}:${address.col}`,
        );
      }
      return { name: font.name, size: font.size, source: effective.source };
    });
    return fonts;
  }, sourceAddresses);
  expect(fonts.map((font) => font.size)).toEqual([15, 17, 19]);
  // EffectiveStyleSource ordinals from the native engine are Default=3,
  // Column=2, and Row=1. These blank cells deliberately exercise each
  // inherited axis without a direct cell XF masking its provenance.
  expect(fonts.map((font) => font.source)).toEqual([3, 2, 1]);
  return fonts;
};

const assertClipboardInheritedFont = (
  source: NativeFont,
  snapshot: { fontFamily: string; fontSize: number },
  pasted: NativeFont,
  label: string,
): void => {
  expect(snapshot.fontFamily, `${label} snapshot family`).toBe(source.name);
  expect(snapshot.fontSize, `${label} snapshot size`).toBeCloseTo(source.size * (96 / 72), 8);
  expect(pasted.name, `${label} pasted family`).toBe(source.name);
  expect(pasted.size, `${label} pasted size`).toBeCloseTo(source.size, 8);
};

const cellCorner = async (
  page: Page,
  row: number,
  col: number,
): Promise<{ x: number; y: number }> => {
  const moduleUrl = `/@fs${new URL('../../../packages/formulon-cell/src/render/geometry.ts', import.meta.url).pathname}`;
  return page.evaluate(
    async ({ moduleUrl: nextModuleUrl, row: targetRow, col: targetCol }) => {
      const instance = (window as ProbeWindow).__fcInst;
      if (!instance) throw new Error('full-check geometry requires __fcInst');
      const { cellRectUnclamped } = await import(nextModuleUrl);
      const { layout, viewport } = instance.store.getState() as unknown as {
        layout: Parameters<typeof cellRectUnclamped>[0];
        viewport: Parameters<typeof cellRectUnclamped>[1];
      };
      const rect = cellRectUnclamped(layout, viewport, targetRow, targetCol);
      const canvas = document.querySelector<HTMLCanvasElement>('.fc-host__canvas');
      const box = canvas?.getBoundingClientRect();
      if (!box) throw new Error('full-check geometry cannot find canvas');
      return { x: box.x + rect.x + 2, y: box.y + rect.y + 2 };
    },
    { moduleUrl, row, col },
  );
};

const errorTriangleCenter = async (
  page: Page,
  row: number,
  col: number,
): Promise<{ x: number; y: number }> => {
  const moduleUrl = `/@fs${
    new URL('../../../packages/formulon-cell/src/render/grid/hit-state.ts', import.meta.url)
      .pathname
  }`;
  return page.evaluate(
    async ({ moduleUrl: nextModuleUrl, row: targetRow, col: targetCol }) => {
      const { gridHitState } = await import(nextModuleUrl);
      const grid = document.querySelector<HTMLElement>('.fc-host__grid');
      if (!grid) throw new Error('full-check error geometry cannot find grid');
      const hit = (gridHitState(grid).errorTriangles as ErrorTriangleProbe[]).find(
        ({ addr }) => addr.sheet === 0 && addr.row === targetRow && addr.col === targetCol,
      );
      if (!hit) {
        throw new Error(`full-check error indicator missing at ${targetRow}:${targetCol}`);
      }
      const canvas = document.querySelector<HTMLCanvasElement>('.fc-host__canvas');
      const box = canvas?.getBoundingClientRect();
      if (!box) throw new Error('full-check error geometry cannot find canvas');
      return {
        x: box.x + hit.rect.x + hit.rect.w / 2,
        y: box.y + hit.rect.y + hit.rect.h / 2,
      };
    },
    { moduleUrl, row, col },
  );
};

const assertRuntime = async (
  page: Page,
  row: FullCheckRow,
  expectedLocale: FullCheckLocale = row.locale,
): Promise<SurfaceLabels> => {
  const labels = await getStrings(page);
  const profile = `${row.platform === 'mac' ? 'mac' : 'windows'}-${PROFILE_TOKENS[expectedLocale]}`;
  await expect(page.locator('.fc-host').first()).toHaveAttribute('lang', expectedLocale);
  await expect(page.locator('.fc-host').first()).toHaveAttribute('data-fc-platform', row.platform);
  await expect(page.locator('.fc-host').first()).toHaveAttribute('data-fc-theme', row.theme);
  await expect(page.locator('html')).toHaveAttribute('lang', expectedLocale);
  await expect.poll(async () => (await runtime(page)).locale).toBe(expectedLocale);
  await expect.poll(async () => (await runtime(page)).profile).toBe(profile);
  await expect(page.locator('[data-ribbon-tab="home"]').first()).toHaveText(labels.home);
  await expect(page.locator('.fc-tb__ribbon-shell').first()).toBeVisible();
  const toolbarTheme = await page
    .locator('.fc-tb__ribbon-shell')
    .first()
    .evaluate((node) => node.closest<HTMLElement>('[data-fc-theme]')?.dataset.fcTheme ?? null);
  expect(toolbarTheme, 'toolbar must inherit the active demo theme').toBe(row.theme);
  return labels;
};

const chooseThemeThroughDemoPanel = async (page: Page, theme: FullCheckTheme): Promise<void> => {
  const toggle = page.locator('.demo__account .demo__share[aria-pressed]').first();
  await toggle.click();
  const panel = page.locator('.demo__panel').first();
  await expect(panel).toBeVisible();
  const themes = panel.locator('.demo__card').first().locator('.demo__seg button');
  await expect(themes).toHaveCount(3);
  const index = { paper: 0, ink: 1, contrast: 2 }[theme];
  await themes.nth(index).click();
  await expect.poll(() => page.locator('.demo').getAttribute('data-fc-theme')).toBe(theme);
  await expect
    .poll(() => page.locator('.fc-host').first().getAttribute('data-fc-theme'))
    .toBe(theme);
  await toggle.click();
  await expect(panel).toBeHidden();
};

const switchLocaleThroughDemoPanel = async (page: Page, locale: FullCheckLocale): Promise<void> => {
  const toggle = page.locator('.demo__account .demo__share[aria-pressed]').first();
  await toggle.click();
  const panel = page.locator('.demo__panel').first();
  await expect(panel).toBeVisible();
  const selector = panel.getByRole('combobox');
  await selector.click();
  await panel.getByRole('option', { name: DEMO_LOCALE_LABELS[locale], exact: true }).click();
  await expect(selector).toHaveAttribute('data-value', locale);
  await expect.poll(async () => (await runtime(page)).locale).toBe(locale);
  await toggle.click();
  await expect(panel).toBeHidden();
};

const attemptRejectedLocaleSwitchWithDraft = async (
  page: Page,
  currentLocale: FullCheckLocale,
  rejectedLocale: FullCheckLocale,
): Promise<void> => {
  const result = await page.evaluate((nextLocale) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check locale draft requires __fcInst');
    return { accepted: instance.setLocale(nextLocale), locale: instance.locale };
  }, rejectedLocale);
  expect(result).toEqual({ accepted: false, locale: currentLocale });
};

const withinViewport = async (
  page: Page,
  locator: import('@playwright/test').Locator,
): Promise<void> => {
  const box = await locator.boundingBox();
  expect(box, 'full-check surface must have geometry').not.toBeNull();
  if (!box) return;
  const viewport = page.viewportSize();
  if (!viewport) throw new Error('full-check viewport is unavailable');
  expect(box.x).toBeGreaterThanOrEqual(-2);
  expect(box.y).toBeGreaterThanOrEqual(-2);
  expect(box.x + box.width).toBeLessThanOrEqual(viewport.width + 2);
  expect(box.y + box.height).toBeLessThanOrEqual(viewport.height + 2);
};

const seedObjectFixture = async (page: Page): Promise<void> => {
  const cells: SeedCell[] = [
    { address: { sheet: 0, row: 0, col: 0 }, kind: 'text', value: 'Region' },
    { address: { sheet: 0, row: 0, col: 1 }, kind: 'text', value: 'Sales' },
    { address: { sheet: 0, row: 1, col: 0 }, kind: 'text', value: 'East' },
    { address: { sheet: 0, row: 1, col: 1 }, kind: 'number', value: 12 },
    { address: { sheet: 0, row: 2, col: 0 }, kind: 'text', value: '' },
    { address: { sheet: 0, row: 2, col: 1 }, kind: 'number', value: 8 },
    { address: { sheet: 0, row: 3, col: 0 }, kind: 'text', value: 'West' },
    { address: { sheet: 0, row: 3, col: 1 }, kind: 'number', value: 20 },
    { address: { sheet: 0, row: 4, col: 0 }, kind: 'text', value: 'North' },
    { address: { sheet: 0, row: 4, col: 1 }, kind: 'number', value: 15 },
  ];
  await seed(page, cells, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
  const formatTableUrl = `/@fs${new URL('../../../packages/formulon-cell/src/commands/format-as-table.ts', import.meta.url).pathname}`;
  const pivotUrl = `/@fs${new URL('../../../packages/formulon-cell/src/commands/pivot-table.ts', import.meta.url).pathname}`;
  const chartUrl = `/@fs${new URL('../../../packages/formulon-cell/src/commands/session-chart.ts', import.meta.url).pathname}`;
  await page.evaluate(
    async ({ formatTableUrl: tableUrl, pivotUrl: nextPivotUrl, chartUrl: nextChartUrl }) => {
      const instance = (window as ProbeWindow).__fcInst;
      if (!instance) throw new Error('full-check objects fixture requires __fcInst');
      if (!instance.workbook.capabilities.tableMutate) {
        throw new Error(
          'full-check objects fixture unavailable: engine tableMutate capability is false',
        );
      }
      if (!instance.workbook.capabilities.pivotTableMutate) {
        throw new Error(
          'full-check objects fixture unavailable: engine pivotTableMutate capability is false',
        );
      }
      instance.setFeatures({ slicer: true });
      const { formatAsTable } = await import(tableUrl);
      const tableRange = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 };
      formatAsTable(instance.store as never, tableRange, {
        workbook: instance.workbook as never,
        style: 'medium',
      });
      const table = instance.workbook.getTables?.()[0];
      if (!table)
        throw new Error('full-check objects fixture failed: table creation returned no table');
      const { createPivotTableFromRange } = await import(nextPivotUrl);
      const pivot = createPivotTableFromRange(instance.workbook as never, {
        source: tableRange,
        destination: { sheet: 0, row: 7, col: 0 },
        name: 'FullCheckPivot',
        rowField: 'Region',
        valueField: 'Sales',
        aggregation: 0,
      });
      if (!pivot.ok) {
        throw new Error(`full-check objects fixture pivot creation failed: ${pivot.reason}`);
      }
      const { createSessionChart } = await import(nextChartUrl);
      const chart = createSessionChart(
        instance.store as never,
        tableRange,
        { wb: instance.workbook as never, kind: 'column' },
        instance.history as never,
      );
      if (!chart)
        throw new Error('full-check objects fixture failed: chart creation returned null');
      const slicer = instance.addSlicer({ tableName: table.name, column: 'Region', x: 24, y: 24 });
      if (!slicer?.id)
        throw new Error('full-check objects fixture failed: slicer creation returned no id');
    },
    { formatTableUrl, pivotUrl, chartUrl },
  );
  await expect.poll(() => page.locator('.fc-chart').count()).toBeGreaterThan(0);
  await expect.poll(() => page.locator('.fc-slicer').count()).toBeGreaterThan(0);
};

const runObjectsFlow = async (page: Page, row: FullCheckRow): Promise<void> => {
  await seedObjectFixture(page);
  await setSelection(page, { sheet: 0, r0: 7, c0: 0, r1: 7, c1: 0 });
  const opened = await page.evaluate(() => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check objects flow requires __fcInst');
    return instance.openActivePivotFieldList();
  });
  expect(opened).toBe(true);
  const fieldList = page.locator('.fc-objects__pivot-field-list').first();
  await expect(fieldList).toBeVisible();
  let labels = await getStrings(page);
  await expect(page.locator('.fc-objects__pivot-field-list-title').first()).toHaveText(
    labels.pivotFieldList,
  );
  await expect(
    fieldList.locator('.fc-objects__pivot-field-list-item').filter({ hasText: 'Region' }),
  ).toBeVisible();
  await withinViewport(page, fieldList);
  await page.keyboard.press('Escape');
  await expect(fieldList).toBeHidden();

  const slicer = page.locator('.fc-slicer').first();
  const chart = page.locator('.fc-chart').first();
  await expect(slicer).toHaveAttribute('aria-label', `${labels.slicerTitle}: Region`);
  await expect(slicer.locator('.fc-slicer__clear')).toHaveText(labels.slicerClear);
  await expect(slicer.locator('.fc-slicer__close')).toHaveAttribute(
    'aria-label',
    labels.slicerClose,
  );
  await expect(slicer.locator('.fc-slicer__chip[data-fc-value=""]').first()).toHaveText(
    labels.blank,
  );
  // One untitled series takes the series name as its automatic title.
  await expect(chart.locator('[data-chart-role="title"]')).toHaveText('Sales');
  const chartElements = chart.locator('.fc-chart__button[data-chart-settings-tab="elements"]');
  await expect(chartElements).toHaveAttribute('aria-label', labels.chartElements);
  await withinViewport(page, slicer);
  await withinViewport(page, chart);

  const nextLocale = localeAfter(row.locale);
  await switchLocaleThroughDemoPanel(page, nextLocale);
  labels = await assertRuntime(page, { ...row, locale: nextLocale }, nextLocale);
  await expect(slicer).toHaveAttribute('aria-label', `${labels.slicerTitle}: Region`);
  await expect(slicer.locator('.fc-slicer__chip[data-fc-value=""]').first()).toHaveText(
    labels.blank,
  );
  await expect(chartElements).toHaveAttribute('aria-label', labels.chartElements);
  await page.keyboard.press('Escape');
  await slicer.locator('.fc-slicer__close').click();
  // A selected chart is deleted with the Delete key; it has no close button.
  await chart.click();
  await page.keyboard.press('Delete');
  await expect(page.locator('.fc-slicer')).toHaveCount(0);
  await expect(page.locator('.fc-chart')).toHaveCount(0);
};

type ClipboardFixture = {
  sp: UserJourneyPage;
  sourceFonts: NativeFont[];
};

const seedClipboardFixture = async (page: Page): Promise<ClipboardFixture> => {
  const sp = new UserJourneyPage(page);
  await seed(
    page,
    [
      { address: { sheet: 0, row: 0, col: 0 }, kind: 'text', value: 'copy-text' },
      { address: { sheet: 0, row: 0, col: 1 }, kind: 'number', value: 42 },
      { address: { sheet: 0, row: 1, col: 0 }, kind: 'formula', value: '=B1*2' },
      { address: { sheet: 0, row: 1, col: 1 }, kind: 'number', value: 21 },
    ],
    { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
  );
  const sourceFonts = await setNativeInheritedClipboardFonts(page);
  // Keep both real values and blank inherited-style cells in one copied range:
  // A1:B2 proves data survives the clipboard path, while D1/E1/D2 prove the
  // workbook, column, and row font axes are copied without direct cell XF
  // materialisation.
  await sp.goTo('A1:E2');
  await sp.shortcut('c');
  await expect
    .poll(() =>
      page.evaluate(() => (window as ProbeWindow).__fcInst?.store.getState().ui.copyMode ?? null),
    )
    .toBe('copy');
  return { sp, sourceFonts };
};

const runClipboardFlow = async (page: Page, row: FullCheckRow): Promise<void> => {
  const { sp, sourceFonts } = await seedClipboardFixture(page);
  const snapshotFonts = await readClipboardSnapshotFonts(page, [
    { row: 0, col: 3 },
    { row: 0, col: 4 },
    { row: 1, col: 3 },
  ]);
  const sourceLabels = ['workbook-default', 'column', 'row'];
  const sourceAddresses: CellAddress[] = [
    { sheet: 0, row: 0, col: 3 },
    { sheet: 0, row: 0, col: 4 },
    { sheet: 0, row: 1, col: 3 },
  ];
  const sourceNativeFonts = await readNativeFonts(page, sourceAddresses);
  expect(sourceNativeFonts).toEqual(sourceFonts);
  expect(sourceNativeFonts.map((font) => font.size)).toEqual([15, 17, 19]);
  // EffectiveStyleSource.Default=3, Column=2, Row=1 in the native API.
  expect(sourceNativeFonts.map((font) => font.source)).toEqual([3, 2, 1]);
  for (let index = 0; index < sourceLabels.length; index += 1) {
    const sourceFont = sourceFonts[index];
    const snapshotFont = snapshotFonts[index];
    const sourceLabel = sourceLabels[index];
    if (!sourceFont || !snapshotFont || !sourceLabel) {
      throw new Error('full-check clipboard fixture produced an incomplete inherited-font matrix');
    }
    // Check the copy payload before any paste mutates the destination. This
    // proves that copy captured the effective native axis style.
    assertClipboardInheritedFont(sourceFont, snapshotFont, sourceFont, `${sourceLabel} source`);
  }
  const addresses: CellAddress[] = [
    { sheet: 0, row: 0, col: 0 },
    { sheet: 0, row: 0, col: 1 },
    { sheet: 0, row: 1, col: 0 },
    { sheet: 0, row: 1, col: 1 },
  ];
  const before = await readCells(page, addresses);
  await page.evaluate(() => (window as ProbeWindow).__fcInst?.openPasteSpecial());
  let labels = await getStrings(page);
  const pasteDialog = page.locator('.fc-pastesp:visible').first();
  await expect(pasteDialog).toBeVisible();
  await expect(pasteDialog).toHaveAttribute('aria-label', labels.pasteSpecialTitle);
  // Six paste-content choices and five arithmetic operations are rendered by
  // the live Paste Special dialog.
  await expect(pasteDialog.locator('input[type="radio"]')).toHaveCount(11);
  await withinViewport(page, pasteDialog);
  const nextLocale = localeAfter(row.locale);
  await attemptRejectedLocaleSwitchWithDraft(page, row.locale, nextLocale);
  await expect(pasteDialog).toHaveAttribute('aria-label', labels.pasteSpecialTitle);
  await pasteDialog.getByRole('button', { name: labels.pasteCancel, exact: true }).click();
  await expect(pasteDialog).toBeHidden();
  expect(await readCells(page, addresses)).toEqual(before);

  await switchLocaleThroughDemoPanel(page, nextLocale);
  labels = await assertRuntime(page, { ...row, locale: nextLocale }, nextLocale);
  // Locale rebinding recreates the engine clipboard adapter. Seed a fresh
  // copy payload before reopening paste dialogs so their assertions exercise
  // the switched profile's live snapshot and labels.
  await sp.goTo('A1:E2');
  await sp.shortcut('c');
  await expect
    .poll(() =>
      page.evaluate(() => (window as ProbeWindow).__fcInst?.store.getState().ui.copyMode ?? null),
    )
    .toBe('copy');
  await page.evaluate(() => (window as ProbeWindow).__fcInst?.openPasteSpecial());
  await expect(pasteDialog).toBeVisible();
  await expect(pasteDialog).toHaveAttribute('aria-label', labels.pasteSpecialTitle);
  await pasteDialog.getByRole('button', { name: labels.pasteCancel, exact: true }).click();
  await expect(pasteDialog).toBeHidden();

  await page.evaluate(() => (window as ProbeWindow).__fcInst?.openInsertCopiedCells());
  const insertDialog = page.locator('.fc-insertcopied:visible').first();
  await expect(insertDialog).toBeVisible();
  await expect(insertDialog).toHaveAttribute('aria-label', labels.insertCopiedTitle);
  await expect(insertDialog.locator('input[type="radio"]')).toHaveCount(2);
  await withinViewport(page, insertDialog);
  await insertDialog.getByRole('button', { name: labels.insertCancel, exact: true }).click();
  await expect(insertDialog).toBeHidden();
  expect(await readCells(page, addresses)).toEqual(before);

  await sp.goTo('H1');
  await sp.shortcut('v');
  await expect
    .poll(() =>
      readCells(page, [
        { sheet: 0, row: 0, col: 7 },
        { sheet: 0, row: 0, col: 8 },
        { sheet: 0, row: 1, col: 7 },
        { sheet: 0, row: 1, col: 8 },
      ]),
    )
    .toEqual([
      { kind: 'text', value: 'copy-text' },
      { kind: 'number', value: 42 },
      { kind: 'number', value: 84 },
      { kind: 'number', value: 21 },
    ]);
  const pastedAddresses: CellAddress[] = [
    { sheet: 0, row: 0, col: 10 },
    { sheet: 0, row: 0, col: 11 },
    { sheet: 0, row: 1, col: 10 },
  ];
  await expect
    .poll(async () => (await readNativeFonts(page, pastedAddresses)).map((font) => font.size))
    .toEqual(sourceFonts.slice(0, 3).map((font) => font.size));
  const pastedFonts = await readNativeFonts(page, pastedAddresses);
  for (let index = 0; index < sourceLabels.length; index += 1) {
    const sourceFont = sourceFonts[index];
    const pastedFont = pastedFonts[index];
    const snapshotFont = snapshotFonts[index];
    const sourceLabel = sourceLabels[index];
    if (!sourceFont || !pastedFont || !snapshotFont || !sourceLabel) {
      throw new Error('full-check clipboard paste produced an incomplete inherited-font matrix');
    }
    assertClipboardInheritedFont(sourceFont, snapshotFont, pastedFont, `${sourceLabel} paste`);
  }
  const pasteOptionsButton = page.locator('.fc-paste-options__button').first();
  await expect(pasteOptionsButton).toBeVisible();
  await pasteOptionsButton.click();
  const pasteOptions = page.locator('.fc-paste-options__menu').first();
  await expect(pasteOptions).toBeVisible();
  await expect(pasteOptions).toHaveAttribute('aria-label', labels.pasteOptionsTitle);
  await expect(pasteOptions).toContainText(labels.pasteKeepSource);
  await expect(pasteOptions).toContainText(labels.pasteValues);
  await expect(pasteOptions).toContainText(labels.pasteFormatting);
  await pasteOptions.getByRole('menuitem').filter({ hasText: labels.pasteValues }).click();
  await expect(pasteOptions).toBeHidden();
  // Pasting a range keeps the whole range selected, so the formula bar is
  // intentionally blank. Navigate to the pasted source cell before checking
  // the cell-level display value.
  await sp.goTo('H1');
  await expect.poll(() => sp.formulaBarValue()).toBe('copy-text');
};

const runFilterFlow = async (page: Page, row: FullCheckRow): Promise<void> => {
  const range = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 2 };
  await seed(
    page,
    [
      { address: { sheet: 0, row: 0, col: 0 }, kind: 'text', value: 'Region' },
      { address: { sheet: 0, row: 0, col: 1 }, kind: 'text', value: 'Sales' },
      { address: { sheet: 0, row: 0, col: 2 }, kind: 'text', value: 'Ready' },
      { address: { sheet: 0, row: 1, col: 0 }, kind: 'text', value: 'East' },
      { address: { sheet: 0, row: 1, col: 1 }, kind: 'number', value: 12.5 },
      { address: { sheet: 0, row: 1, col: 2 }, kind: 'bool', value: true },
      { address: { sheet: 0, row: 2, col: 0 }, kind: 'text', value: '' },
      { address: { sheet: 0, row: 2, col: 1 }, kind: 'number', value: 8 },
      { address: { sheet: 0, row: 2, col: 2 }, kind: 'bool', value: false },
      { address: { sheet: 0, row: 3, col: 0 }, kind: 'text', value: 'West' },
      { address: { sheet: 0, row: 3, col: 1 }, kind: 'number', value: 20 },
      { address: { sheet: 0, row: 3, col: 2 }, kind: 'bool', value: true },
      { address: { sheet: 0, row: 4, col: 0 }, kind: 'text', value: 'North' },
      { address: { sheet: 0, row: 4, col: 1 }, kind: 'number', value: 15 },
      { address: { sheet: 0, row: 4, col: 2 }, kind: 'bool', value: false },
      { address: { sheet: 0, row: 5, col: 2 }, kind: 'bool', value: true },
    ],
    range,
  );
  const autoFilterUrl = `/@fs${new URL('../../../packages/formulon-cell/src/engine/auto-filter-sync.ts', import.meta.url).pathname}`;
  await page.evaluate(
    async ({ moduleUrl, nextRange }) => {
      const instance = (window as ProbeWindow).__fcInst;
      if (!instance) throw new Error('full-check filter fixture requires __fcInst');
      if (!instance.workbook.capabilities.autoFilter || !instance.workbook.setSheetAutoFilterXml) {
        throw new Error(
          'full-check filter fixture unavailable: engine autoFilter capability is false',
        );
      }
      if (!instance.workbook.setSheetAutoFilterXml(0, '<autoFilter ref="A1:C6"/>')) {
        throw new Error('full-check filter fixture failed: engine rejected autoFilter definition');
      }
      const { hydrateAutoFilterFromEngine } = await import(moduleUrl);
      hydrateAutoFilterFromEngine(instance.workbook as never, instance.store as never, 0);
      instance.store.setState((state) => ({
        ...state,
        selection: {
          active: { sheet: 0, row: 0, col: 0 },
          anchor: { sheet: 0, row: 0, col: 0 },
          range: nextRange,
          extraRanges: [],
        },
      }));
    },
    { moduleUrl: autoFilterUrl, nextRange: range },
  );
  const openFilter = async (): Promise<void> => {
    await page.evaluate((nextRange) => {
      const instance = (window as ProbeWindow).__fcInst;
      if (!instance) throw new Error('full-check filter flow requires __fcInst');
      instance.openFilterDropdown(nextRange, 1);
    }, range);
  };
  await openFilter();
  const dropdown = page.locator('.fc-filter-dropdown:visible').first();
  await expect(dropdown).toBeVisible();
  let labels = await getStrings(page);
  await expect(dropdown).toHaveAttribute('aria-label', labels.filterTitle);
  await expect(dropdown.locator('.fc-filter-dropdown__search')).toHaveAttribute(
    'placeholder',
    labels.filterSearch,
  );
  await expect(dropdown).toContainText(labels.filterSelectAll);
  await expect(dropdown).toContainText(labels.filterBlanks);
  await expect(dropdown).toContainText(localizedDecimal(row.locale, 12.5));
  await expect(dropdown.locator('.fc-filter-dropdown__row')).toHaveCount(6);
  await withinViewport(page, dropdown);
  const nextLocale = localeAfter(row.locale);
  const beforeLocaleSwitch = await runtime(page);
  // A floating filter is a live surface: changing the profile through the
  // host options must refresh its labels in place while preserving its
  // current range, criteria, and row visibility.
  await switchLocaleThroughDemoPanel(page, nextLocale);
  labels = await assertRuntime(page, { ...row, locale: nextLocale }, nextLocale);
  // The options panel rebuilds transient popups; verify the accepted locale on
  // a freshly opened filter surface before applying the condition.
  await openFilter();
  await expect(dropdown).toBeVisible();
  await expect(dropdown).toHaveAttribute('aria-label', labels.filterTitle);
  await expect(dropdown.locator('.fc-filter-dropdown__search')).toHaveAttribute(
    'placeholder',
    labels.filterSearch,
  );
  await expect(dropdown).toContainText(labels.filterSelectAll);
  await expect(dropdown).toContainText(labels.filterBlanks);
  await expect(dropdown).toContainText(labels.filterApply);
  await expect(dropdown).toContainText(labels.filterClear);
  await expect(dropdown).toContainText(localizedDecimal(nextLocale, 12.5));
  await expect(dropdown.locator('.fc-filter-dropdown__row')).toHaveCount(6);
  const afterLocaleSwitch = await runtime(page);
  expect(afterLocaleSwitch.filterRanges).toEqual(beforeLocaleSwitch.filterRanges);
  expect(afterLocaleSwitch.filterCriteria).toEqual(beforeLocaleSwitch.filterCriteria);
  expect(afterLocaleSwitch.hiddenRows).toEqual(beforeLocaleSwitch.hiddenRows);
  expect(
    await readCells(page, [
      { sheet: 0, row: 1, col: 1 },
      { sheet: 0, row: 2, col: 1 },
      { sheet: 0, row: 3, col: 1 },
    ]),
  ).toEqual([
    { kind: 'number', value: 12.5 },
    { kind: 'number', value: 8 },
    { kind: 'number', value: 20 },
  ]);
  await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();
  await openFilter();
  await expect(dropdown).toBeVisible();
  await dropdown.locator('.fc-filter-dropdown__condition-op').selectOption('greaterThan');
  await dropdown
    .locator('.fc-filter-dropdown__condition-value')
    .fill(localizedDecimal(nextLocale, 10));
  await dropdown.locator('.fc-filter-dropdown__apply').click();
  await expect(dropdown).toBeHidden();
  await expect.poll(async () => (await runtime(page)).filterRanges).toEqual([range]);
  await expect.poll(async () => (await runtime(page)).hiddenRows).toEqual([2, 5]);
  await expect
    .poll(async () => (await runtime(page)).filterCriteria)
    .toEqual([
      {
        range,
        byCol: 1,
        hiddenValues: [],
        condition: { op: 'greaterThan', value: '10' },
      },
    ]);
  await openFilter();
  await expect(dropdown).toBeVisible();
  await dropdown.locator('.fc-filter-dropdown__clear').click();
  await expect(dropdown).toBeHidden();
  await expect.poll(async () => (await runtime(page)).hiddenRows).toEqual([]);
  await expect.poll(async () => (await runtime(page)).filterCriteria).toEqual([]);
  await expect.poll(async () => (await runtime(page)).filterRanges).toEqual([]);

  await page.evaluate((nextRange) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check boolean filter requires __fcInst');
    instance.openFilterDropdown(nextRange, 2);
  }, range);
  await expect(dropdown).toBeVisible();
  await expect(dropdown).toContainText(localizedBoolean(nextLocale, true));
  await expect(dropdown).toContainText(localizedBoolean(nextLocale, false));
  await expect(dropdown.locator('.fc-filter-dropdown__row')).toHaveCount(3);
  await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();
};

const localizedFormula = async (page: Page, canonical: string): Promise<string> =>
  page.evaluate((formula) => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check formula requires __fcInst');
    const workbook = instance.workbook as typeof instance.workbook & {
      localizeFormulaText?: (value: string) => string;
    };
    return workbook.localizeFormulaText?.(formula) ?? formula;
  }, canonical);

const runFormulaFlow = async (page: Page, row: FullCheckRow): Promise<void> => {
  await seed(
    page,
    [
      { address: { sheet: 0, row: 0, col: 0 }, kind: 'number', value: 1 },
      { address: { sheet: 0, row: 1, col: 0 }, kind: 'number', value: 2 },
    ],
    { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
  );
  const sp = new UserJourneyPage(page);
  const canonicalFormula = '=SUM(A1:A2)';
  const expectedLocalizedFormula = (locale: FullCheckLocale): string => {
    if (locale === 'de') return '=SUMME(A1:A2)';
    if (locale === 'fr') return '=SOMME(A1:A2)';
    return canonicalFormula;
  };
  const completeFormula = await localizedFormula(page, canonicalFormula);
  const localizedSumName = completeFormula.slice(1, completeFormula.indexOf('('));
  const localizedSum = `=${localizedSumName}(`;
  // The autocomplete resolver matches the identifier before the caret. Keep
  // the opening parenthesis for the separate argument-helper assertion below.
  const autocompletePrefix = completeFormula.slice(0, completeFormula.indexOf('('));
  await sp.goTo('C1');
  await page.keyboard.type(autocompletePrefix);
  const editor = page.locator('.fc-host__editor').first();
  await expect(editor).toBeVisible();
  const autocomplete = page.locator('.fc-autocomplete').first();
  await expect(autocomplete).toBeVisible();
  const labels = await getStrings(page);
  await expect(editor).toHaveAttribute('aria-expanded', 'true');
  await expect
    .poll(() => autocomplete.locator('.fc-autocomplete__item').count())
    .toBeGreaterThan(0);
  await expect(autocomplete).toContainText(
    new RegExp(`${labels.autocompletePick}|${labels.autocompleteCustom}`),
  );
  await page.keyboard.press('Escape');
  await expect(autocomplete).toBeHidden();

  // Keep the argument helper assertion separate from the committed formula:
  // Escape must discard this draft rather than creating a cell formula.
  await sp.goTo('D1');
  await page.keyboard.type(localizedSum);
  const argHelper = page.locator('.fc-arghelper').first();
  await expect(argHelper).toBeVisible();
  await expect(argHelper.locator('.fc-arghelper__name').first()).toHaveText(`${localizedSumName}(`);
  await expect(argHelper.locator('.fc-arghelper__sep').first()).toHaveText(
    row.locale === 'de' || row.locale === 'fr' ? '; ' : ', ',
  );
  await expect(argHelper.locator('.fc-arghelper__arg[aria-current="true"]')).toHaveText('number1');
  await page.keyboard.press('Escape');
  await expect(argHelper).toBeHidden();
  expect(await readCellFormula(page, { sheet: 0, row: 0, col: 3 })).toBeNull();

  expect(completeFormula).toBe(expectedLocalizedFormula(row.locale));
  await sp.goTo('C1');
  await page.keyboard.type(completeFormula);
  await page.keyboard.press('Enter');
  await expect
    .poll(() => readCellFormula(page, { sheet: 0, row: 0, col: 2 }))
    .toBe(canonicalFormula);
  await expect
    .poll(() => readCells(page, [{ sheet: 0, row: 0, col: 2 }]))
    .toEqual([{ kind: 'number', value: 3 }]);

  // The next locale is deliberately chosen from a profile with a translated
  // SUM alias. en/ja/zh/ko/th retain canonical English function names, so
  // simply taking the next matrix locale would fail to prove a visible change.
  const nextLocale: FullCheckLocale = row.locale === 'de' ? 'fr' : 'de';
  await switchLocaleThroughDemoPanel(page, nextLocale);
  await assertRuntime(page, { ...row, locale: nextLocale }, nextLocale);
  const nextFormula = await localizedFormula(page, canonicalFormula);
  expect(nextFormula).toBe(expectedLocalizedFormula(nextLocale));
  await sp.goTo('C1');
  await expect.poll(() => sp.formulaBarValue()).toBe(nextFormula);
  expect(await readCellFormula(page, { sheet: 0, row: 0, col: 2 })).toBe(canonicalFormula);
  await expect
    .poll(() => readCells(page, [{ sheet: 0, row: 0, col: 2 }]))
    .toEqual([{ kind: 'number', value: 3 }]);

  // Reopen a fresh draft under the switched profile to prove the argument
  // helper follows the translated formula surface as well.
  await sp.goTo('D1');
  const nextLocalizedSumName = nextFormula.slice(1, nextFormula.indexOf('('));
  await page.keyboard.type(`=${nextLocalizedSumName}(`);
  const nextArgHelper = page.locator('.fc-arghelper:visible').first();
  await expect(nextArgHelper).toBeVisible();
  await expect(nextArgHelper.locator('.fc-arghelper__name').first()).toHaveText(
    `${nextLocalizedSumName}(`,
  );
  await expect(nextArgHelper.locator('.fc-arghelper__sep').first()).toHaveText('; ');
  await expect(nextArgHelper.locator('.fc-arghelper__arg[aria-current="true"]')).toHaveText(
    'number1',
  );
  await page.keyboard.press('Escape');
};

const runAnalysisFlow = async (page: Page, row: FullCheckRow): Promise<void> => {
  await page.evaluate(() => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check analysis feature setup requires __fcInst');
    instance.setFeatures({ errorIndicators: true });
  });
  const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 };
  await seed(
    page,
    [
      { address: { sheet: 0, row: 0, col: 0 }, kind: 'number', value: 1 },
      { address: { sheet: 0, row: 0, col: 1 }, kind: 'number', value: 4 },
      { address: { sheet: 0, row: 1, col: 0 }, kind: 'number', value: 2 },
      { address: { sheet: 0, row: 1, col: 1 }, kind: 'number', value: 5 },
      { address: { sheet: 0, row: 2, col: 0 }, kind: 'number', value: 3 },
      { address: { sheet: 0, row: 2, col: 1 }, kind: 'number', value: 6 },
      { address: { sheet: 0, row: 5, col: 0 }, kind: 'formula', value: '=1/0' },
    ],
    range,
  );
  const openQuickAnalysis = async (): Promise<void> => {
    await page.evaluate(() => {
      const instance = (window as ProbeWindow).__fcInst;
      if (!instance) throw new Error('full-check analysis flow requires __fcInst');
      instance.openQuickAnalysis();
    });
  };
  await openQuickAnalysis();
  // Locale changes rebuild the transient sheet. Ignore the detached old root
  // and always follow the visible, newly rendered quick-analysis surface.
  const quick = page.locator('.fc-quick:visible').first();
  await expect(quick).toBeVisible();
  let labels = await getStrings(page);
  await expect(quick.locator('.fc-quick__title')).toHaveText(labels.quickTitle);
  await expect(quick.locator('.fc-quick__tab')).toHaveCount(5);
  await expect(quick.locator('[data-action="format-data-bar"]')).toHaveText(labels.quickDataBar);
  await withinViewport(page, quick);
  await switchLocaleThroughDemoPanel(page, localeAfter(row.locale));
  labels = await assertRuntime(
    page,
    { ...row, locale: localeAfter(row.locale) },
    localeAfter(row.locale),
  );
  await openQuickAnalysis();
  await expect(quick.locator('.fc-quick__title')).toHaveText(labels.quickTitle);
  await page.keyboard.press('Escape');
  await expect(quick).toBeHidden();
  await setSelection(page, range);
  await openQuickAnalysis();
  await expect(quick).toBeVisible();
  const ruleCount = (await runtime(page)).conditionalRuleCount;
  await quick.locator('[data-action="format-data-bar"]').click();
  await expect
    .poll(async () => (await runtime(page)).conditionalRuleCount)
    .toBeGreaterThan(ruleCount);
  const undone = await page.evaluate(() => {
    const instance = (window as ProbeWindow).__fcInst;
    if (!instance) throw new Error('full-check analysis undo requires __fcInst');
    return (instance as BrowserInstance & { undo(): boolean }).undo();
  });
  expect(undone).toBe(true);
  await expect.poll(async () => (await runtime(page)).conditionalRuleCount).toBe(ruleCount);

  const errorPoint = await errorTriangleCenter(page, 5, 0);
  const dragStart = await cellCorner(page, 0, 1);
  const ribbonBox = await page.locator('.fc-tb__ribbon-shell').first().boundingBox();
  if (!ribbonBox) throw new Error('full-check error pointer regression cannot find ribbon');
  const outsidePoint = {
    x: ribbonBox.x + ribbonBox.width / 2,
    y: ribbonBox.y + ribbonBox.height / 2,
  };
  // A drag that starts in another cell and ends over the error marker must not
  // open the error menu: the marker is actionable only for a click that began
  // on the marker itself.
  await page.mouse.move(dragStart.x, dragStart.y);
  await page.mouse.down();
  await page.mouse.move(errorPoint.x, errorPoint.y, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator('.fc-errmenu:visible')).toHaveCount(0);

  // Leaving the marker while pressed disarms the hit. Returning before
  // release must remain a drag and still leave the menu closed.
  await page.mouse.move(errorPoint.x, errorPoint.y);
  await page.mouse.down();
  await page.mouse.move(outsidePoint.x, outsidePoint.y, { steps: 5 });
  await page.mouse.move(errorPoint.x, errorPoint.y, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator('.fc-errmenu:visible')).toHaveCount(0);

  // The final physical click proves the legitimate marker path still opens
  // the localized menu after the drag guards have reset.
  await page.mouse.click(errorPoint.x, errorPoint.y);
  const errorMenu = page.locator('.fc-errmenu:visible').first();
  await expect(errorMenu).toBeVisible();
  labels = await getStrings(page);
  await expect(errorMenu).toHaveAttribute('aria-label', labels.errorHeading);
  await expect(errorMenu.locator('[data-fc-action="showInfo"]')).toHaveText(labels.errorShowInfo);
  await expect(errorMenu.locator('[data-fc-action="ignore"]')).toHaveText(labels.errorIgnore);
  await page.keyboard.press('Escape');
  await expect(errorMenu).toBeHidden();
  await page.mouse.click(errorPoint.x, errorPoint.y);
  await expect(errorMenu).toBeVisible();
  await errorMenu.locator('[data-fc-action="ignore"]').click();
  await expect(errorMenu).toBeHidden();
  await expect.poll(async () => (await runtime(page)).ignoredErrorCount).toBe(1);
};

const runMenusFlow = async (page: Page, row: FullCheckRow): Promise<void> => {
  const sp = new UserJourneyPage(page);
  await seed(page, [{ address: { sheet: 0, row: 0, col: 0 }, kind: 'text', value: 'menu-copy' }], {
    sheet: 0,
    r0: 0,
    c0: 0,
    r1: 0,
    c1: 0,
  });
  await sp.goTo('A1');
  await sp.shortcut('c');
  await expect
    .poll(() =>
      page.evaluate(() => (window as ProbeWindow).__fcInst?.store.getState().ui.copyMode ?? null),
    )
    .toBe('copy');
  let point = await cellCorner(page, 0, 0);
  await page.mouse.click(point.x + 20, point.y + 20, { button: 'right' });
  const contextMenu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)').first();
  await expect(contextMenu).toBeVisible();
  let labels = await getStrings(page);
  await expect(contextMenu).toHaveAttribute('aria-label', labels.contextTitle);
  await expect(contextMenu).toContainText(labels.copy);
  await expect(contextMenu).toContainText(labels.paste);
  await expect(contextMenu).toContainText(labels.pasteSpecial);
  await expect(contextMenu).toContainText(labels.insertCopiedCells);
  await page.keyboard.press('Escape');
  await expect(contextMenu).toBeHidden();
  await switchLocaleThroughDemoPanel(page, localeAfter(row.locale));
  labels = await assertRuntime(
    page,
    { ...row, locale: localeAfter(row.locale) },
    localeAfter(row.locale),
  );
  await expect
    .poll(() =>
      page.evaluate(() => (window as ProbeWindow).__fcInst?.store.getState().ui.copyMode ?? null),
    )
    .toBe('copy');
  point = await cellCorner(page, 0, 0);
  await page.mouse.click(point.x + 20, point.y + 20, { button: 'right' });
  await expect(contextMenu).toBeVisible();
  await expect(contextMenu).toHaveAttribute('aria-label', labels.contextTitle);
  await expect(contextMenu).toContainText(labels.copy);
  await page.keyboard.press('Escape');

  const fileTab = page.locator('[data-ribbon-tab="file"]').first();
  if (await fileTab.count()) {
    await fileTab.click();
  } else {
    await page.locator('.demo__brand-mark').first().click();
  }
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await expect(backstage).toHaveAttribute('aria-label', labels.backstageFile);
  await expect(
    backstage.getByRole('button', { name: labels.backstageInfo, exact: true }),
  ).toBeVisible();
  await expect(
    backstage.getByRole('button', { name: labels.backstageClose, exact: true }),
  ).toBeVisible();
  await withinViewport(page, backstage);
  await backstage.getByRole('button', { name: labels.backstageClose, exact: true }).click();
  await expect(backstage).toBeHidden();

  const readDisplayMode = async (): Promise<{ shell: string | null; toolbar: string | null }> =>
    page.evaluate(() => {
      const instance = (window as ProbeWindow).__fcInst;
      const shell = document.querySelector<HTMLElement>('.fc-tb__ribbon-shell');
      if (!instance || !shell) {
        throw new Error('full-check display mode requires the live ribbon instance');
      }
      return {
        shell: shell.dataset.ribbonDisplayMode ?? null,
        toolbar:
          instance.toolbar?.getDisplayMode?.() ??
          (window as ProbeWindow).__fcToolbar?.getDisplayMode?.() ??
          null,
      };
    });

  await expect.poll(readDisplayMode).toEqual({ shell: 'full', toolbar: 'full' });
  const displayToggle = page.locator('[data-ribbon-toggle]').first();
  await displayToggle.click();
  const displayMenu = page.locator('.fc-tb__ribbon-display-menu').first();
  await expect(displayMenu).toBeVisible();
  labels = await getStrings(page);
  await expect(displayToggle).toHaveAttribute('aria-label', labels.displayLabel);
  for (const text of [
    labels.displayExpanded,
    labels.displaySingleLine,
    labels.displayCollapsed,
    labels.displayAutoHide,
  ]) {
    await expect(displayMenu).toContainText(text);
  }
  await displayMenu.locator('[data-ribbon-display-option="singleLine"]').click();
  await expect.poll(readDisplayMode).toEqual({ shell: 'singleLine', toolbar: 'singleLine' });
  await displayToggle.click();
  await expect(displayMenu).toBeVisible();
  await expect(displayMenu.locator('[data-ribbon-display-option="singleLine"]')).toHaveAttribute(
    'aria-checked',
    'true',
  );
  await displayMenu.locator('[data-ribbon-display-option="full"]').click();
  await expect.poll(readDisplayMode).toEqual({ shell: 'full', toolbar: 'full' });
  await expect(displayMenu).toBeHidden();
};

export async function runFullCheckMatrixRow(page: Page, row: FullCheckRow): Promise<void> {
  const spreadsheet = new SpreadsheetPage(page);
  const errors = spreadsheet.collectConsoleErrors();
  await page.setViewportSize({ width: Number(row.width), height: 800 });
  await spreadsheet.mount({ fixture: 'empty', locale: row.locale, platform: row.platform });
  await spreadsheet.expectNoStub();
  await chooseThemeThroughDemoPanel(page, row.theme);
  await assertRuntime(page, row);

  switch (row.flow) {
    case 'objects':
      await runObjectsFlow(page, row);
      break;
    case 'clipboard':
      await runClipboardFlow(page, row);
      break;
    case 'filter':
      await runFilterFlow(page, row);
      break;
    case 'formula':
      await runFormulaFlow(page, row);
      break;
    case 'analysis':
      await runAnalysisFlow(page, row);
      break;
    case 'menus':
      await runMenusFlow(page, row);
      break;
  }

  expect(
    errors.read(),
    `${row.locale}/${row.platform}/${row.theme}/${row.width}/${row.flow} console errors`,
  ).toEqual([]);
}
