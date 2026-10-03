import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import {
  RIBBON_BORDERS_MENU_ID,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
} from '../../../../src/toolbar/ribbon/activation.js';
import {
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS,
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS,
  DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS,
  type DynamicDropdownsCtx,
} from '../../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import { ribbonDir, root, sourceFilesUnder, sourcesOutsidePrimitives } from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
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
    expect(dynamicDropdownsSource.match(/applyVerticalViewportLimit\(/g) ?? []).toHaveLength(3);
    expect(dynamicDropdownsSource).toMatch(
      /applyVerticalViewportLimit\(panel, null, 0\);\s+panel\.hidden = false;/,
    );
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
});
