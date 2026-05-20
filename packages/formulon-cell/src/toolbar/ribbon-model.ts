import { dictionaries, type Strings } from '../i18n/strings.js';
import { pageScaleMenuText, toolbarMenuText, viewToggleMenuText } from './menu-text.js';

export type RibbonTab =
  | 'file'
  | 'home'
  | 'insert'
  | 'draw'
  | 'pageLayout'
  | 'formulas'
  | 'data'
  | 'review'
  | 'view'
  | 'help'
  | 'automate'
  | 'acrobat';

export type ToolbarLang = 'ja' | 'en';

export interface RibbonCommand {
  id: string;
  title: string;
  label: string;
  icon?: string;
  kind?: 'button' | 'large' | 'wide' | 'mono' | 'select' | 'color' | 'break';
  layout?: 'stacked';
  disabled?: boolean;
  options?: readonly RibbonOption[];
  className?: string;
}

export interface RibbonOption {
  value: string;
  label: string;
}

export interface RibbonGroupModel {
  title: string;
  variant?: string;
  commands: RibbonCommand[];
}

export interface RibbonTabModel {
  id: RibbonTab;
  label: string;
  groups: RibbonGroupModel[];
}

export const HOME_TILE_LAYOUT_GROUP_VARIANTS = ['styles'] as const;
export const HOME_STACKED_LAYOUT_GROUP_VARIANTS = ['cells'] as const;
export const HOME_MIXED_LAYOUT_GROUP_VARIANTS = ['editing'] as const;

export const EXCEL365_STANDARD_RIBBON_TABS: readonly RibbonTab[] = [
  'file',
  'home',
  'insert',
  'pageLayout',
  'formulas',
  'data',
  'review',
  'view',
  'help',
];

export const OPTIONAL_RIBBON_TABS: readonly RibbonTab[] = ['draw', 'automate', 'acrobat'];

export const RIBBON_TABS: readonly RibbonTab[] = [
  ...EXCEL365_STANDARD_RIBBON_TABS,
  ...OPTIONAL_RIBBON_TABS,
];

export const RIBBON_TAB_LABELS: Readonly<Record<RibbonTab, Record<ToolbarLang, string>>> = {
  file: { ja: 'ファイル', en: 'File' },
  home: { ja: 'ホーム', en: 'Home' },
  insert: { ja: '挿入', en: 'Insert' },
  draw: { ja: '描画', en: 'Draw' },
  pageLayout: { ja: 'ページ レイアウト', en: 'Page Layout' },
  formulas: { ja: '数式', en: 'Formulas' },
  data: { ja: 'データ', en: 'Data' },
  review: { ja: '校閲', en: 'Review' },
  view: { ja: '表示', en: 'View' },
  help: { ja: 'ヘルプ', en: 'Help' },
  automate: { ja: '自動化', en: 'Automate' },
  acrobat: { ja: 'Acrobat', en: 'Acrobat' },
};

export const RIBBON_KEYSHORTCUTS: Readonly<Record<string, string>> = {
  copy: 'Control+C Meta+C',
  cut: 'Control+X Meta+X',
  findHome: 'Control+F Meta+F',
  findReview: 'Control+F Meta+F',
  formatCells: 'Control+1 Meta+1',
  formatCellsHome: 'Control+1 Meta+1',
  fx: 'Shift+F3',
  fxInsert: 'Shift+F3',
  gotoSpecial: 'F5 Control+G Meta+G',
  gotoSpecialHome: 'F5 Control+G Meta+G',
  hyperlinkInsert: 'Control+K Meta+K',
  namedRanges: 'Control+F3',
  paste: 'Control+V Meta+V',
  recalcNow: 'F9',
  redoHome: 'Control+Y Meta+Y Meta+Shift+Z',
  undoHome: 'Control+Z Meta+Z',
};

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36] as const;

export const FONT_FAMILIES = [
  'Aptos',
  'Aptos Display',
  'Aptos Narrow',
  'Calibri',
  'Arial',
  'Segoe UI',
  '游ゴシック Light',
  '游ゴシック Regular',
  'Yu Gothic UI',
  'BIZ UDGothic',
  'BIZ UDMincho',
  'Times New Roman',
  'Consolas',
  'Georgia',
  'Tahoma',
  'Verdana',
] as const;

export type ToolbarText = Strings['ribbon'];

const resolveStrings = (input: Strings | ToolbarLang): Strings =>
  typeof input === 'string' ? dictionaries[input] : input;

export const toolbarText = (input: Strings | ToolbarLang): ToolbarText =>
  resolveStrings(input).ribbon;

export const ribbonTabLabel = (input: Strings | ToolbarLang, id: RibbonTab): string =>
  resolveStrings(input).ribbon.tabs[id];

export interface BuildRibbonModelOptions {
  /** Explicit tab surface. Use `EXCEL365_STANDARD_RIBBON_TABS` for the
   *  Microsoft 365 baseline and append `OPTIONAL_RIBBON_TABS` only when the
   *  host really exposes those add-in/automation surfaces. */
  tabs?: readonly RibbonTab[];
}

const COMMAND_SURFACE_LANG: ToolbarLang = 'en';

const interpolate = (template: string, vars: Record<string, string>): string =>
  template.replace(/\{(\w+)\}/g, (_, key) => vars[key] ?? '');

const cmd = (
  id: string,
  label: string,
  title = label,
  icon?: string,
  kind: RibbonCommand['kind'] = 'button',
  disabled = false,
): RibbonCommand => ({ id, title, label, icon, kind, disabled });

const selectCmd = (
  id: string,
  label: string,
  title: string,
  options: readonly RibbonOption[],
  className?: string,
): RibbonCommand => ({ id, title, label, kind: 'select', options, className });

const colorCmd = (id: string, label: string, title: string, icon: string): RibbonCommand => ({
  id,
  title,
  label,
  icon,
  kind: 'color',
});

const breakCmd = (id: string): RibbonCommand => ({ id, title: '', label: '', kind: 'break' });

export function buildRibbonModel(
  input: Strings | ToolbarLang,
  opts: BuildRibbonModelOptions = {},
): RibbonTabModel[] {
  const strings = resolveStrings(input);
  const tr = strings.ribbon;
  const menuText = toolbarMenuText(strings);
  const pageScaleText = pageScaleMenuText(strings);
  const viewText = viewToggleMenuText(strings);
  const scalePageOption = (value: '1' | '2' | '3'): string =>
    `${value} ${value === '1' ? pageScaleText.page : pageScaleText.pages}`;
  const functionArgsTitle = (name: string): string =>
    interpolate(tr.functionArgumentsTitle, { name });
  const zoomTitle = (percent: string): string => interpolate(tr.zoomToTitle, { percent });
  const commentTitle = (kind: 'delete' | 'previous' | 'next'): string => {
    if (kind === 'delete') return tr.deleteCommentOrNoteTitle;
    if (kind === 'previous') return tr.previousCommentOrNoteTitle;
    return tr.nextCommentOrNoteTitle;
  };
  const outlineTitle = (kind: 'group' | 'ungroup' | 'show' | 'hide'): string => {
    if (kind === 'group') return tr.groupSelectedRowsOrColumnsTitle;
    if (kind === 'ungroup') return tr.ungroupSelectedRowsOrColumnsTitle;
    if (kind === 'show') return tr.showGroupedDetailTitle;
    return tr.hideGroupedDetailTitle;
  };
  const tab = (id: RibbonTab, groups: RibbonGroupModel[]): RibbonTabModel => ({
    id,
    label: tr.tabs[id],
    groups,
  });
  const group = (
    title: string,
    commands: RibbonCommand[],
    variant: RibbonGroupModel['variant'] = 'tiles',
  ): RibbonGroupModel => ({ title, commands, variant });
  const allTabs = [
    tab('file', [
      group(tr.workbook, [
        cmd('pageSetup', tr.pageSetup, tr.pageSetup, 'page', 'wide'),
        cmd('print', tr.print, tr.print, 'print', 'wide'),
        cmd('protect', tr.protect, tr.protect, 'protect', 'wide'),
      ]),
      group(tr.inspect, [cmd('inspect', tr.inspect, tr.inspect, 'goTo', 'wide')]),
    ]),
    tab('home', [
      group(
        tr.clipboard,
        [
          cmd('paste', tr.paste, tr.paste, 'paste', 'large'),
          cmd('cut', tr.cut, tr.cut, 'cut'),
          cmd('copy', tr.copy, tr.copy, 'copy'),
          cmd('formatPainter', tr.formatPainter, tr.formatPainter, 'paint'),
        ],
        'clipboard',
      ),
      group(
        tr.font,
        [
          selectCmd(
            'fontFamily',
            tr.font,
            tr.font,
            tr.fontFamilies.map((font) => ({ value: font, label: font })),
            'demo__rb-select--font',
          ),
          selectCmd(
            'fontSize',
            tr.fontSize,
            tr.fontSize,
            FONT_SIZES.map((size) => ({ value: String(size), label: String(size) })),
          ),
          cmd('fontGrow', '', tr.increaseFontSize, 'fontGrow'),
          cmd('fontShrink', '', tr.decreaseFontSize, 'fontShrink'),
          breakCmd('font-row-2'),
          cmd('bold', 'B', `${tr.bold} (⌘B)`, 'bold'),
          cmd('italic', 'I', `${tr.italic} (⌘I)`, 'italic'),
          cmd('underline', 'U', `${tr.underline} (⌘U)`, 'underline'),
          cmd('strike', 'S', tr.strikethrough, 'strike'),
          cmd('borders', tr.borders, tr.borders, 'borders'),
          colorCmd('fillColor', tr.fillColor, tr.fillColor, 'fillColor'),
          colorCmd('fontColor', tr.fontColor, tr.fontColor, 'fontColor'),
        ],
        'font',
      ),
      group(
        tr.alignment,
        [
          cmd('top', tr.top, tr.topAlign, 'top'),
          cmd('middle', tr.middle, tr.middleAlign, 'middle'),
          cmd('bottomAlign', tr.bottomAlign, tr.bottomAlign, 'bottomAlign'),
          cmd('alignL', tr.alignLeft, tr.alignLeft, 'alignLeft'),
          cmd('alignC', tr.alignCenter, tr.alignCenter, 'alignCenter'),
          cmd('alignR', tr.alignRight, tr.alignRight, 'alignRight'),
          breakCmd('alignment-row-2'),
          cmd('textOrientation', tr.textOrientation, tr.textOrientation, 'textOrientation'),
          cmd('wrap', tr.wrapText, tr.wrapText, 'wrap'),
          cmd('indentDecrease', tr.decreaseIndent, tr.decreaseIndent, 'indentDecrease'),
          cmd('indentIncrease', tr.increaseIndent, tr.increaseIndent, 'indentIncrease'),
          cmd('merge', tr.mergeCells, tr.mergeCells, 'merge'),
        ],
        'alignment',
      ),
      group(
        tr.number,
        [
          selectCmd(
            'numberFormat',
            tr.general,
            tr.number,
            [
              { value: 'general', label: tr.general },
              { value: 'fixed', label: tr.fixedNumber },
              { value: 'currency', label: tr.currency },
              { value: 'accounting', label: tr.accounting },
              { value: 'shortDate', label: tr.shortDate },
              { value: 'longDate', label: tr.longDate },
              { value: 'time', label: tr.timeFormat },
              { value: 'percent', label: tr.percent },
              { value: 'fraction', label: tr.fraction },
              { value: 'scientific', label: tr.scientific },
              { value: 'text', label: tr.textFormat },
              { value: 'more', label: tr.moreNumberFormats },
            ],
            'demo__rb-select--number-format',
          ),
          breakCmd('number-row-2'),
          cmd('currency', tr.currency, tr.currency, 'currency'),
          cmd('percent', '%', tr.percent, 'percent', 'mono'),
          cmd('comma', ',', tr.commaStyle, 'comma'),
          cmd('decDown', '.0', tr.decreaseDecimals, 'decDown', 'mono'),
          cmd('decUp', '.00', tr.increaseDecimals, 'decUp', 'mono'),
        ],
        'number',
      ),
      group(
        tr.styles,
        [
          cmd('conditional', tr.conditional, tr.conditionalFormatting, 'conditional', 'wide'),
          cmd('formatTableHome', tr.formatTable, tr.formatTable, 'tableStyle', 'wide'),
          cmd('cellStyles', tr.cellStyles, tr.cellStyles, 'tableStyle', 'wide'),
        ],
        'styles',
      ),
      group(
        tr.cells,
        [
          {
            ...cmd('insertRows', tr.insert, tr.insertRows, 'insertRows', 'wide'),
            layout: 'stacked',
          },
          {
            ...cmd('deleteRows', tr.delete, tr.deleteRows, 'deleteRows', 'wide'),
            layout: 'stacked',
          },
          {
            ...cmd('formatCellsHome', tr.format, tr.formatCells, 'formatCells', 'wide'),
            layout: 'stacked',
          },
        ],
        'cells',
      ),
      group(
        tr.editing,
        [
          {
            ...cmd('autosum', tr.autoSum, `${tr.autoSum} (Σ)`, 'autosum', 'wide'),
            layout: 'stacked',
          },
          {
            ...cmd('fillHome', tr.fill, tr.fill, 'fillColor', 'wide'),
            layout: 'stacked',
          },
          {
            ...cmd('clearFormat', tr.clear, tr.clear, 'clear', 'wide'),
            layout: 'stacked',
          },
          cmd('sortFilterHome', tr.sortFilter, tr.sortFilter, 'sortAsc', 'wide'),
          cmd('findHome', tr.findSelect, `${tr.findSelect} (⌘F)`, 'find', 'wide'),
        ],
        'editing',
      ),
    ]),
    tab('insert', [
      group(tr.tables, [
        cmd('pivotTableInsert', tr.pivotTable, tr.pivotTable, 'table', 'wide'),
        cmd('formatTableInsert', tr.table, tr.table, 'table', 'wide'),
      ]),
      group(tr.illustrations, [
        cmd('pictureInsert', tr.pictures, tr.pictures, 'page', 'wide'),
        cmd('shapesInsert', tr.shapes, tr.shapes, 'scale', 'wide'),
        cmd('screenshotInsert', tr.screenshot, tr.screenshot, 'goTo', 'wide'),
      ]),
      group(tr.charts, [cmd('chartInsert', tr.chart, tr.chart, 'chart', 'wide')]),
      group(tr.links, [
        cmd('hyperlinkInsert', tr.hyperlink, `${tr.hyperlink} (⌘K)`, 'link', 'wide'),
      ]),
      group(tr.comments, [
        cmd('commentInsert', tr.newComment, tr.newComment, 'commentAdd', 'wide'),
      ]),
      group(tr.symbols, [cmd('symbolInsert', tr.symbol, tr.symbol, 'function', 'wide')]),
    ]),
    tab('pageLayout', [
      group(menuText.theme, [cmd('pageTheme', menuText.theme, menuText.theme, 'options', 'wide')]),
      group(tr.pageSetup, [
        selectCmd(
          'marginsPreset',
          tr.margins,
          tr.margins,
          [
            { value: 'normal', label: tr.marginsNormal },
            { value: 'wide', label: tr.marginsWide },
            { value: 'narrow', label: tr.marginsNarrow },
            { value: 'custom', label: tr.marginsCustom },
          ],
          'demo__rb-select--margins',
        ),
        selectCmd(
          'orientationPreset',
          tr.orientation,
          tr.orientation,
          [
            { value: 'portrait', label: tr.portrait },
            { value: 'landscape', label: tr.landscape },
          ],
          'demo__rb-select--border',
        ),
        selectCmd(
          'paperSizePreset',
          tr.paperSize,
          tr.paperSize,
          [
            { value: 'A4', label: tr.paperA4 },
            { value: 'A3', label: tr.paperA3 },
            { value: 'A5', label: tr.paperA5 },
            { value: 'letter', label: tr.paperLetter },
            { value: 'legal', label: tr.paperLegal },
            { value: 'tabloid', label: tr.paperTabloid },
          ],
          'demo__rb-select--border',
        ),
        cmd('pageSetupAdvanced', tr.pageSetup, tr.pageSetup, 'options', 'wide'),
        cmd(
          'printArea',
          tr.printArea,
          `${tr.printArea}: ${menuText.printAreaSet}/${menuText.printAreaAdd}/${menuText.printAreaClear}`,
          'table',
          'wide',
        ),
        cmd('pageBreaks', tr.breaks, tr.breaks, 'page', 'wide'),
        cmd('sheetBackground', tr.background, tr.background, 'page', 'wide'),
        cmd('printTitles', tr.printTitles, tr.printTitles, 'table', 'wide'),
      ]),
      group(tr.scale, [
        selectCmd(
          'scaleWidth',
          pageScaleText.width,
          pageScaleText.fitWidth,
          [
            { value: '0', label: pageScaleText.automatic },
            { value: '1', label: scalePageOption('1') },
            { value: '2', label: scalePageOption('2') },
            { value: '3', label: scalePageOption('3') },
            { value: 'custom', label: pageScaleText.custom },
          ],
          'demo__rb-select--border',
        ),
        selectCmd(
          'scaleHeight',
          pageScaleText.height,
          pageScaleText.fitHeight,
          [
            { value: '0', label: pageScaleText.automatic },
            { value: '1', label: scalePageOption('1') },
            { value: '2', label: scalePageOption('2') },
            { value: '3', label: scalePageOption('3') },
            { value: 'custom', label: pageScaleText.custom },
          ],
          'demo__rb-select--border',
        ),
        selectCmd(
          'scalePercent',
          pageScaleText.scale,
          tr.scale,
          [
            { value: '25', label: '25%' },
            { value: '50', label: '50%' },
            { value: '75', label: '75%' },
            { value: '100', label: '100%' },
            { value: '125', label: '125%' },
            { value: '150', label: '150%' },
            { value: '200', label: '200%' },
            { value: '400', label: '400%' },
            { value: 'custom', label: pageScaleText.custom },
          ],
          'demo__rb-select--border',
        ),
      ]),
      group(tr.sheetOptions, [
        cmd(
          'pageLayoutGridlinesView',
          `${viewText.gridlines} ${tr.show}`,
          viewText.gridlines,
          'table',
          'wide',
        ),
        cmd(
          'pageLayoutGridlinesPrint',
          strings.pageSetup.showGridlines,
          strings.pageSetup.showGridlines,
          'print',
          'wide',
        ),
        cmd(
          'pageLayoutHeadingsView',
          `${viewText.headings} ${tr.show}`,
          viewText.headings,
          'table',
          'wide',
        ),
        cmd(
          'pageLayoutHeadingsPrint',
          strings.pageSetup.showHeadings,
          strings.pageSetup.showHeadings,
          'print',
          'wide',
        ),
      ]),
      group(tr.arrange, [
        cmd('arrangeObjectsPageLayout', tr.arrange, tr.arrange, 'options', 'wide'),
        cmd('selectionPanePageLayout', tr.selectionPane, tr.selectionPane, 'options', 'wide'),
      ]),
      group(tr.print, [cmd('printPageLayout', tr.print, tr.print, 'print', 'wide')]),
    ]),
    tab('formulas', [
      group(tr.functionLibrary, [
        cmd('fx', 'fx', tr.insertFunction, 'function', 'mono'),
        cmd('autosumFormula', tr.autoSum, `${tr.autoSum} (Σ)`, 'autosum', 'wide'),
        cmd('sum', 'SUM', functionArgsTitle('SUM'), 'function', 'mono'),
        cmd('avg', 'AVG', functionArgsTitle('AVERAGE'), 'function', 'mono'),
        cmd('ifFormula', 'IF', functionArgsTitle('IF'), 'function', 'mono'),
        cmd('xlookupFormula', 'XLOOKUP', functionArgsTitle('XLOOKUP'), 'function', 'mono'),
        cmd('concatFormula', 'CONCAT', functionArgsTitle('CONCAT'), 'function', 'mono'),
        cmd('todayFormula', 'TODAY', functionArgsTitle('TODAY'), 'function', 'mono'),
        cmd('pmtFormula', 'PMT', functionArgsTitle('PMT'), 'function', 'mono'),
        cmd('roundFormula', 'ROUND', functionArgsTitle('ROUND'), 'function', 'mono'),
      ]),
      group(tr.definedNames, [cmd('namedRanges', tr.names, tr.names, 'names', 'wide')]),
      group(tr.formulaAuditing, [
        cmd('precedents', tr.tracePrecedents, tr.tracePrecedents, 'trace', 'wide'),
        cmd('dependents', tr.traceDependents, tr.traceDependents, 'dependents', 'wide'),
        cmd('clearArrows', tr.removeArrows, tr.removeArrows, 'clearArrows', 'wide'),
        cmd('errorChecking', tr.errorChecking, tr.errorChecking, 'options', 'wide'),
        cmd('showFormulasFormula', viewText.formulas, viewText.formulas, 'function', 'wide'),
        cmd('evaluateFormula', tr.evaluateFormula, tr.evaluateFormula, 'function', 'wide'),
      ]),
      group(tr.calculation, [
        cmd('recalcNow', tr.recalc, `${tr.recalc} (F9)`, 'autosum', 'wide'),
        cmd('calcOptions', tr.options, tr.options, 'options', 'wide'),
        cmd('watch', tr.watch, tr.watch, 'watch', 'wide'),
      ]),
    ]),
    tab('data', [
      group(tr.sortFilter, [
        cmd('filter', tr.filter, tr.filter, 'filter', 'wide'),
        cmd('sortAsc', 'A-Z', tr.sortAscending, 'sortAsc'),
        cmd('sortDesc', 'Z-A', tr.sortDescending, 'sortDesc'),
        cmd('sortData', menuText.sortCustom, menuText.sortCustom, 'sortAsc', 'wide'),
      ]),
      group(tr.dataTools, [
        cmd('textToColumns', menuText.textToColumns, menuText.textToColumns, 'table', 'wide'),
        cmd('removeDupes', tr.removeDuplicates, tr.removeDuplicates, 'removeDuplicates', 'wide'),
        cmd('dataValidation', tr.dataValidation, tr.dataValidation, 'options', 'wide'),
        cmd('linksData', tr.links, tr.links, 'link', 'wide'),
      ]),
      group(tr.outline, [
        cmd('outlineGroup', tr.groupOutline, outlineTitle('group'), 'table', 'wide'),
        cmd('outlineUngroup', tr.ungroupOutline, outlineTitle('ungroup'), 'table', 'wide'),
        cmd('outlineShowDetail', tr.showDetail, outlineTitle('show'), 'table', 'wide'),
        cmd('outlineHideDetail', tr.hideDetail, outlineTitle('hide'), 'table', 'wide'),
      ]),
    ]),
    tab('review', [
      group(tr.proofing, [cmd('spellingReview', tr.spelling, tr.spelling, 'spelling', 'wide')]),
      group(tr.accessibility, [
        cmd('accessibility', tr.accessibility, tr.accessibility, 'accessibility', 'wide'),
      ]),
      group(tr.language, [cmd('translateReview', tr.translate, tr.translate, 'translate', 'wide')]),
      group(tr.comments, [
        cmd('newCommentReview', tr.newComment, tr.newComment, 'commentAdd', 'wide'),
        cmd('deleteCommentReview', tr.deleteComment, commentTitle('delete'), 'clear', 'wide'),
        cmd('previousCommentReview', tr.previousComment, commentTitle('previous'), 'goTo', 'wide'),
        cmd('nextCommentReview', tr.nextComment, commentTitle('next'), 'goTo', 'wide'),
      ]),
      group(tr.find, [cmd('findReview', tr.find, `${tr.find} (⌘F)`, 'find', 'wide')]),
      group(tr.protection, [
        cmd('protectReview', tr.protect, tr.protect, 'protect', 'wide'),
        cmd(
          'protectWorkbookReview',
          menuText.protectWorkbookCommand,
          menuText.protectWorkbookCommand,
          'protect',
          'wide',
        ),
        cmd(
          'protectionReview',
          menuText.allowEditRangesCommand,
          menuText.allowEditRangesCommand,
          'protect',
          'wide',
        ),
      ]),
    ]),
    tab('view', [
      group(tr.workbookViews, [
        cmd('viewNormal', tr.normalView, tr.normalView, 'table', 'wide'),
        cmd('viewPageLayout', tr.pageLayoutView, tr.pageLayoutView, 'page', 'wide'),
        cmd('viewPageBreakPreview', tr.pageBreakPreview, tr.pageBreakPreview, 'table', 'wide'),
        cmd('watchView', tr.watch, tr.watch, 'watch', 'wide'),
      ]),
      group(strings.viewToolbar.views, [
        selectCmd(
          'sheetViewSelect',
          strings.viewToolbar.currentView,
          strings.viewToolbar.views,
          [{ value: 'current', label: strings.viewToolbar.currentView }],
          'demo__rb-select--border',
        ),
        cmd('sheetViewSave', strings.viewToolbar.saveView, strings.viewToolbar.saveView, 'options'),
        cmd(
          'sheetViewDelete',
          strings.viewToolbar.deleteView,
          strings.viewToolbar.deleteView,
          'clear',
        ),
        cmd(
          'workbookObjectsView',
          strings.viewToolbar.objects,
          strings.viewToolbar.objects,
          'options',
        ),
        cmd(
          'pivotFieldListView',
          strings.workbookObjects.pivotFieldList,
          strings.workbookObjects.pivotFieldList,
          'options',
        ),
      ]),
      group(tr.show, [
        cmd('viewGridlines', viewText.gridlines, viewText.gridlines, 'table', 'wide'),
        cmd('viewHeadings', viewText.headings, viewText.headings, 'table', 'wide'),
        cmd('viewFormulas', viewText.formulas, viewText.formulas, 'function', 'wide'),
        cmd('viewFormulaBar', viewText.formulaBar, viewText.formulaBar, 'function', 'wide'),
        cmd('viewR1C1', 'R1C1', tr.r1c1, 'options', 'wide'),
      ]),
      group(tr.window, [
        cmd('freeze', tr.freeze, tr.freeze, 'freeze', 'wide'),
        cmd('windowVisibility', tr.format, tr.format, 'table', 'wide'),
      ]),
      group(tr.zoom, [
        cmd('zoomDialog', `${tr.zoom}...`, tr.zoom, 'zoom', 'wide'),
        cmd('zoomSelection', tr.zoomSelection, tr.zoomSelection, 'zoom', 'wide'),
        cmd('zoom75', '75%', zoomTitle('75%'), 'zoom', 'mono'),
        cmd('zoom100', '100%', zoomTitle('100%'), 'zoom', 'mono'),
        cmd('zoom125', '125%', zoomTitle('125%'), 'zoom', 'mono'),
      ]),
      group(tr.protection, [cmd('protect', tr.protect, tr.protect, 'protect', 'wide')]),
    ]),
    tab('help', [
      group(tr.tabs.help, [cmd('helpSearch', tr.tabs.help, tr.tabs.help, 'options', 'wide', true)]),
    ]),
    tab('draw', [
      group(tr.tabs.draw, [
        cmd('drawPen', tr.pen, tr.tabs.draw, 'pen', 'wide'),
        cmd('drawGrid', tr.drawBorderGrid, tr.drawBorderGrid, 'borders', 'wide'),
        cmd('drawErase', tr.eraser, tr.eraser, 'eraser', 'wide'),
      ]),
    ]),
    tab('automate', [
      group(tr.tabs.automate, [
        cmd('script', tr.script, tr.script, 'script', 'wide'),
        cmd('recordActions', tr.recordActions, tr.recordActions, 'script', 'wide'),
        cmd('allScripts', tr.allScripts, tr.allScripts, 'script', 'wide'),
      ]),
    ]),
    tab('acrobat', [
      group(tr.addIn, [cmd('addIn', tr.addIn, tr.addIn, 'addIn', 'wide')]),
      group(tr.pdf, [cmd('pdf', tr.pdf, tr.pdf, 'pdf', 'wide')]),
    ]),
  ];
  const allowed = new Set(opts.tabs ?? RIBBON_TABS);
  return allTabs.filter((tab) => allowed.has(tab.id));
}

export const ribbonCommands = (
  input: Strings | ToolbarLang,
  opts: BuildRibbonModelOptions = {},
): RibbonCommand[] =>
  buildRibbonModel(input, opts)
    .flatMap((tab) => tab.groups)
    .flatMap((group) => group.commands)
    .filter((command) => command.kind !== 'break');

export const isRibbonActivatableCommand = (command: RibbonCommand): boolean =>
  command.kind !== 'break' && !['select', 'color'].includes(command.kind ?? 'button');

export const ribbonActivatableCommands = (
  input: Strings | ToolbarLang,
  opts: BuildRibbonModelOptions = {},
): RibbonCommand[] => ribbonCommands(input, opts).filter(isRibbonActivatableCommand);

export const ribbonCommandIds = (
  input: Strings | ToolbarLang,
  opts: BuildRibbonModelOptions = {},
): string[] => ribbonCommands(input, opts).map((command) => command.id);

export const ribbonActivatableCommandIds = (
  input: Strings | ToolbarLang,
  opts: BuildRibbonModelOptions = {},
): string[] => ribbonActivatableCommands(input, opts).map((command) => command.id);

export const ribbonSurfaceCommands = (
  opts: BuildRibbonModelOptions = {},
): RibbonCommand[] => ribbonCommands(COMMAND_SURFACE_LANG, opts);

export const ribbonSurfaceCommandIds = (opts: BuildRibbonModelOptions = {}): string[] =>
  ribbonCommandIds(COMMAND_SURFACE_LANG, opts);

export const ribbonActivatableSurfaceCommands = (
  opts: BuildRibbonModelOptions = {},
): RibbonCommand[] => ribbonActivatableCommands(COMMAND_SURFACE_LANG, opts);

export const ribbonActivatableSurfaceCommandIds = (
  opts: BuildRibbonModelOptions = {},
): string[] => ribbonActivatableCommandIds(COMMAND_SURFACE_LANG, opts);

export const ribbonTabCommandIds = (
  input: Strings | ToolbarLang,
  tabId: RibbonTab,
  opts: BuildRibbonModelOptions = {},
): string[] =>
  buildRibbonModel(input, opts)
    .find((tab) => tab.id === tabId)
    ?.groups.flatMap((group) => group.commands.map((command) => command.id)) ?? [];
