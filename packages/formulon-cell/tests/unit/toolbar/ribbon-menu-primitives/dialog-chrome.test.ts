import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { root, sourceFilesUnder } from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
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
      'src/interact/format-dialog-tabs/number-controller.ts',
      'src/interact/format-dialog-tabs/font-controller.ts',
      'src/interact/format-dialog-tabs/more.ts',
    ].map((path) => readFileSync(join(root, path), 'utf8'));
    const interactSurfaceSources = [
      'src/interact/filter-dropdown.ts',
      'src/interact/view-toolbar.ts',
      'src/interact/pivot-field-settings.ts',
      'src/interact/pivot-table-dialog.ts',
      'src/interact/workbook-objects-pivot-editor.ts',
      'src/interact/page-setup-dialog.ts',
      'src/interact/page-setup-dialog-tabs/page.ts',
      'src/interact/page-setup-dialog-tabs/header-footer.ts',
      'src/interact/page-setup-dialog-tabs/sheet.ts',
      'src/interact/conditional-form-controls.ts',
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

  it('keeps Format Cells chrome close to Japanese Excel 365 desktop', () => {
    const themeSource = readFileSync(join(root, 'src/styles/theme-paper.css'), 'utf8');
    const frameSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/frame.css'),
      'utf8',
    );
    const controlsSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/controls.css'),
      'utf8',
    );
    const choicesSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/choices.css'),
      'utf8',
    );
    const tabsContentSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/tabs-content.css'),
      'utf8',
    );
    const numberSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/number.css'),
      'utf8',
    );
    const swatchesSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/swatches-and-lines.css'),
      'utf8',
    );
    const bordersSource = readFileSync(
      join(root, 'src/styles/core/app/format-dialog/borders.css'),
      'utf8',
    );
    const customSelectSource = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/custom-select.css'),
      'utf8',
    );

    expect(themeSource).toContain('--fc-fmtdlg-tab-active-bg: #107c41');
    expect(themeSource).toContain('--fc-fmtdlg-tab-active-color: #ffffff');
    expect(themeSource).toContain('--fc-fmtdlg-tab-active-underline: none');
    expect(themeSource).toContain('--fc-fmtdlg-list-hover-bg: #eeeeee');
    expect(frameSource).toContain('border-radius: var(--fc-fmtdlg-tab-radius, 0)');
    expect(frameSource).toContain('filter: none');
    expect(controlsSource).toContain('border-radius: 4px');
    expect(controlsSource).toContain('box-shadow: 0 0 0 1px var(--fc-fmtdlg-list-focus-border)');
    expect(controlsSource).not.toContain('box-shadow: 0 0 0 2px var(--fc-accent-soft)');
    expect(controlsSource).toMatch(
      /\.fc-range-picker__btn::after\s*\{[\s\S]*?width: 8px;[\s\S]*?height: 8px;[\s\S]*?linear-gradient\(45deg,[\s\S]*?#185abd[\s\S]*?content: "";/,
    );
    expect(controlsSource).not.toContain('content: "↗"');
    expect(choicesSource).toContain('border-radius: 4px');
    expect(tabsContentSource).toContain(
      'box-shadow: inset 0 0 0 1px var(--fc-fmtdlg-list-focus-border)',
    );
    expect(tabsContentSource).toContain(
      '.fc-fmtdlg__font-list-item:hover {\n    background: var(--fc-fmtdlg-list-hover-bg);',
    );
    expect(numberSource).toContain(
      'box-shadow: inset 0 0 0 1px var(--fc-fmtdlg-list-focus-border)',
    );
    expect(swatchesSource).toContain('border-radius: 2px');
    expect(swatchesSource).toContain('border: 1px solid var(--fc-fmtdlg-input-hover-border)');
    expect(bordersSource).toContain('border-radius: 2px');
    expect(customSelectSource).toMatch(/\.fc-select__button\s*\{[\s\S]*?border-radius: 2px;/);
    expect(customSelectSource).toContain('box-shadow: none');
    expect(customSelectSource).toMatch(/\.fc-select__list\s*\{[\s\S]*?border-radius: 2px;/);
    expect(customSelectSource).toMatch(/\.fc-select__option\s*\{[\s\S]*?border-radius: 0;/);
    expect(customSelectSource).not.toContain('box-shadow: 0 0 0 2px var(--fc-accent-soft');
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
    expect(shellSource).toContain("button.className = opts.className ?? 'fc-tb__cf-choice'");
    expect(symbolSource).toContain('createDialogChoiceButton({ label: symbol');
    expect(symbolSource).not.toContain("const button = document.createElement('button')");
    expect(symbolSource).not.toContain("button.className = 'fc-tb__cf-choice'");
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
});
