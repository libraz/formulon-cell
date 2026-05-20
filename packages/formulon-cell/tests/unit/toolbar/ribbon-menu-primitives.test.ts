import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, vi } from 'vitest';
import {
  RIBBON_BORDERS_MENU_ID,
  RIBBON_DROPDOWN_MENU_FOR_COMMAND,
} from '../../../src/toolbar/ribbon/activation.js';
import {
  focusMenuItem,
  handleMenuKeydown,
  prepareMenu,
  projectDisabledReason,
  projectDisabledState,
} from '../../../src/toolbar/menu-a11y.js';
import {
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS,
  DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS,
  DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS,
  type DynamicDropdownsCtx,
} from '../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import {
  colorSwatchGrid,
  colorSwatchButton,
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

describe('toolbar/ribbon menu primitives', () => {
  const sourcesOutsidePrimitives = (): { name: string; source: string }[] =>
    menuSources().filter(({ name }) => name !== 'general.ts');

  it('keeps preset menu row DOM centralized in menuPresetButton', () => {
    const directPresetRows = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes('app__menu-item app__menu-item--preset'))
      .map(({ name }) => name);

    expect(directPresetRows).toEqual([]);
  });

  it('keeps iconic menu row DOM centralized in menuIconButton', () => {
    const directIconicRows = sourcesOutsidePrimitives()
      .filter(
        ({ source }) =>
          source.includes('app__menu-item app__menu-item--iconic') ||
          source.includes('app__menu-icon app__menu-icon--'),
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
    expect(backstageTitleSource).not.toContain("ribbonLang === 'ja' ? `「${query}」");
    expect(backstageTitleSource).not.toContain('No matches for "${query}"');
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
    expect(dynamicDefaultsSource).toContain('const invalidRange = strings.advancedFilterInvalidRange');
    expect(dynamicDefaultsSource).toContain('showRenameSheetDialog: (opts) =>');
    expect(dynamicDefaultsSource).toContain('okLabel: strings.hyperlinkDialog.ok');
    expect(dynamicDefaultsSource).toContain('cancelLabel: strings.hyperlinkDialog.cancel');
    expect(dynamicDefaultsSource).toContain('projectFormatToolbar: opts.projectFormatToolbar ?? noop');
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
    expect(promptSource).not.toContain("Enter a valid number.");
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
          /className\s*=\s*['"][^'"]*app__visual-tile/.test(source) ||
          /className\s*=\s*['"][^'"]*app__visual-grid/.test(source) ||
          /className\s*=\s*['"][^'"]*app__color-swatch/.test(source) ||
          /className\s*=\s*['"][^'"]*app__symbol-tile/.test(source) ||
          /className\s*=\s*['"][^'"]*app__symbol-grid/.test(source),
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
      .filter(({ source }) => source.includes("className = 'app__menu-heading'"))
      .map(({ name }) => name);

    expect(directHeadings).toEqual([]);
  });

  it('keeps submenu trigger affordances centralized in menuSubmenuTrigger', () => {
    const directSubmenuTriggers = sourcesOutsidePrimitives()
      .filter(
        ({ source }) =>
          source.includes('app__menu-item--submenu') ||
          source.includes('app__menu-item__caret') ||
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
      "menuSubmenuTrigger(btn, { cfSubmenu: key }, { controlsId: cfSubmenuId(key) })",
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
      .filter(({ source }) => source.includes('app__menu-item__icon-spacer'))
      .map(({ name }) => name);

    expect(directSpacers).toEqual([]);
  });

  it('keeps submenu item text DOM centralized in submenuItemText', () => {
    const directSubmenuText = sourcesOutsidePrimitives()
      .filter(({ source }) => source.includes('app__submenu-item__text'))
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
          /className\s*=\s*['"]app__(?:table|cell)style-heading['"]/.test(source) ||
          /className\s*=\s*['"]app__(?:table|cell)style-grid['"]/.test(source),
      )
      .map(({ name }) => name);

    expect(directLabeledGridDom).toEqual([]);
    expect(new Map(menuSources().map(({ name, source }) => [name, source])).get('styles.ts')).toContain(
      'menuLabeledGrid(',
    );
  });

  it('keeps gallery scroll bodies centralized in menuScrollBody', () => {
    const stylesSource = new Map(menuSources().map(({ name, source }) => [name, source])).get(
      'styles.ts',
    );

    expect(stylesSource).toContain('menuScrollBody(');
    expect(stylesSource).not.toMatch(/className\s*=\s*['"]app__tablestyle-scroll['"]/);
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
      /className\s*=\s*['"]app__cf-(?:choice-row|choice-grid-panel|icon-panel)['"]/,
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
      applyFillSeries: vi.fn(),
      applyFillDirection: vi.fn(),
      applyClearAction: vi.fn(),
      applyUnderlineAction: vi.fn(),
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
    const routedHooks = Array.from(new Set(Object.values(DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS)))
      .sort();

    expect(routedHooks).toEqual(updateHooks);
  });

  it('keeps dynamic dropdown handler dataset keys derived from the shared handler attrs', () => {
    const datasetKeyForAttr = (attr: string): string =>
      attr.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
    const expected = new Set([
      ...DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS.map(datasetKeyForAttr),
      'cfAction',
      'cfSubmenu',
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
      /const DYNAMIC_DROPDOWN_HANDLERS:[\s\S]*?= \[([\s\S]*?)\n  \];/,
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
    expect(dynamicDropdownsSource.split('\n').filter((line) => line.includes('event.target'))).toHaveLength(
      1,
    );
    expect(dynamicDropdownsSource.split('\n').filter((line) => line.includes('.disabled'))).toHaveLength(
      1,
    );
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
    expect(shellSource).toContain("button.className = opts.className ?? 'app__cf-choice'");
    expect(symbolSource).toContain('createDialogChoiceButton({ label: symbol');
    expect(symbolSource).not.toContain("const button = document.createElement('button')");
    expect(symbolSource).not.toContain("button.className = 'app__cf-choice'");
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
      { button: menuPresetButton('Bottom', 'borderPreset', 'bottom', leading), key: 'borderPreset' },
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
          className: 'app__menu-item app__cellstyle-chip',
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
      className: 'app__menu-item app__menu-item--custom',
      attr: 'sampleAction',
      value: 'run',
      title: 'Run sample',
      ariaLabel: 'Run sample',
    });

    expect(button.className).toBe('app__menu-item app__menu-item--custom');
    expect(button.type).toBe('button');
    expect(button.getAttribute('role')).toBe('menuitem');
    expect(button.dataset.sampleAction).toBe('run');
    expect(button.title).toBe('Run sample');
    expect(button.getAttribute('aria-label')).toBe('Run sample');
  });

  it('creates menu div primitives with shared class and accessibility contracts', () => {
    const menu = createMenu('menu-test');
    expect(menu.id).toBe('menu-test');
    expect(menu.className).toBe('app__menu');
    expect(menu.hidden).toBe(true);

    const colorGrid = colorSwatchGrid('test-colors');
    expect(colorGrid.className).toBe('app__color-swatch-grid test-colors');
    expect(colorGrid.getAttribute('role')).toBe('presentation');

    const symbolGrid = symbolMenuGrid('Greek', ['π']);
    expect(symbolGrid.className).toBe('app__symbol-grid');
    expect(symbolGrid.getAttribute('role')).toBe('presentation');
    expect(symbolGrid.getAttribute('aria-label')).toBe('Greek');
    expect(symbolGrid.querySelectorAll('button')).toHaveLength(1);

    const visualGrid = visualMenuGrid('test-visuals');
    expect(visualGrid.className).toBe('app__visual-grid test-visuals');
    expect(visualGrid.getAttribute('role')).toBe('presentation');

    const separator = menuSeparator();
    expect(separator.className).toBe('app__menu-sep');
    expect(separator.getAttribute('role')).toBe('separator');

    const heading = menuSectionHeader('Styles');
    expect(heading.className).toBe('app__menu-heading');
    expect(heading.getAttribute('role')).toBe('presentation');
    expect(heading.textContent).toBe('Styles');

    const [labeledHeading, labeledGrid] = menuLabeledGrid({
      label: 'Light',
      headingClassName: 'app__tablestyle-heading',
      gridClassName: 'app__tablestyle-grid',
      children: [],
    });
    expect(labeledHeading.className).toBe('app__tablestyle-heading');
    expect(labeledHeading.textContent).toBe('Light');
    expect(labeledGrid.className).toBe('app__tablestyle-grid');
    expect(labeledGrid.getAttribute('role')).toBe('group');
    expect(labeledGrid.getAttribute('aria-label')).toBe('Light');
  });

  it('creates nested submenus with the shared menu contract', () => {
    const submenu = createSubmenu({
      id: 'menu-test-submenu',
      className: 'app__submenu app__submenu--test',
      label: 'Test submenu',
      dataset: { cfPanel: 'highlight' },
    });

    expect(submenu.id).toBe('menu-test-submenu');
    expect(submenu.className).toBe('app__submenu app__submenu--test');
    expect(submenu.getAttribute('role')).toBe('menu');
    expect(submenu.getAttribute('aria-label')).toBe('Test submenu');
    expect(submenu.hidden).toBe(true);
    expect(submenu.dataset.cfPanel).toBe('highlight');
  });

  it('decorates submenu triggers with caret and shared accessibility attributes', () => {
    const button = menuPresetButton('Highlight Cells Rules', 'cfAction', 'submenu-highlight', document.createElement('span'));
    const trigger = menuSubmenuTrigger(
      button,
      { cfSubmenu: 'highlight' },
      { controlsId: 'menu-conditional-highlight' },
    );

    expect(trigger).toBe(button);
    expect(trigger.classList.contains('app__menu-item--submenu')).toBe(true);
    expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    expect(trigger.getAttribute('aria-controls')).toBe('menu-conditional-highlight');
    expect(trigger.dataset.cfSubmenu).toBe('highlight');
    const caret = trigger.querySelector<HTMLElement>('.app__menu-item__caret');
    expect(caret?.textContent).toBe('▶');
    expect(caret?.getAttribute('aria-hidden')).toBe('true');
  });

  it('creates shared preset icon spacers', () => {
    const spacer = menuIconSpacer();

    expect(spacer.tagName).toBe('SPAN');
    expect(spacer.className).toBe('app__menu-item__icon-spacer');
  });

  it('creates shared submenu item text labels', () => {
    const text = submenuItemText('None');

    expect(text.tagName).toBe('SPAN');
    expect(text.className).toBe('app__submenu-item__text');
    expect(text.textContent).toBe('None');
  });
});
