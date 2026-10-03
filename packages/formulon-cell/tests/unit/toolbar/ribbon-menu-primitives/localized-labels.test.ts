import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  dynamicDropdownDefaultsSource,
  menusDir,
  mountDir,
  ribbonDir,
  root,
  sourcesOutsidePrimitives,
} from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
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
    const defaultsSource = dynamicDropdownDefaultsSource();

    expect(defaultsSource).toContain('strings.ribbonMenu.cellStyleMergeImported.replace');
    expect(defaultsSource).not.toContain('cellStyleMergeImported?:');
    expect(defaultsSource).not.toContain('style(s) imported');
  });

  it('keeps Create Table dialog labels backed by shared dialog strings', () => {
    const defaultsSource = dynamicDropdownDefaultsSource();

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
    const defaultsSource = dynamicDropdownDefaultsSource();

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
    const dynamicDefaultsSource = dynamicDropdownDefaultsSource();

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
    const dropdownCss = readFileSync(join(root, 'src/styles/toolbar/ribbon/dropdowns.css'), 'utf8');

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
    expect(selectColorSource).not.toContain("arrow.textContent = '›'");
    expect(dropdownCss).toMatch(
      /\.fc-tb__rb-dd__submenu\s*\{[\s\S]*?border-top: 4px solid transparent;[\s\S]*?border-bottom: 4px solid transparent;[\s\S]*?border-left: 5px solid var\(--fc-tb-fg\);/,
    );
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
    expect(backstageTitleSource).not.toContain(`ribbonLang === 'ja' ? \`「\${query}」`);
    expect(backstageTitleSource).not.toContain(`No matches for "\${query}"`);
  });

  it('uses shared localized labels for report dialogs from default toolbar glue', () => {
    const dynamicDefaultsSource = dynamicDropdownDefaultsSource();
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
    const dynamicDefaultsSource = dynamicDropdownDefaultsSource();
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
    expect(dynamicDefaultsSource).toContain(
      'const invalidRange = strings.advancedFilterInvalidRange',
    );
    expect(dynamicDefaultsSource).toContain('showRenameSheetDialog: (opts) =>');
    expect(dynamicDefaultsSource).toContain('okLabel: strings.hyperlinkDialog.ok');
    expect(dynamicDefaultsSource).toContain('cancelLabel: strings.hyperlinkDialog.cancel');
    expect(dynamicDefaultsSource).toContain(
      'projectFormatToolbar: opts.projectFormatToolbar ?? noop',
    );
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
    expect(promptSource).not.toContain('Enter a valid number.');
  });

  it('keeps ribbon menu labels from branching directly on Japanese locale', () => {
    const localeBranchedMenus = sourcesOutsidePrimitives()
      .filter(
        ({ source }) => source.includes("ribbonLang === 'ja' ?") || source.includes('const ja ='),
      )
      .map(({ name }) => name);

    expect(localeBranchedMenus).toEqual([]);
  });
});
