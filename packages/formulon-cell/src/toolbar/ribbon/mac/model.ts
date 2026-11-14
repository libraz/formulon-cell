import { allFunctionNames } from '../../../commands/function-categories.js';
import { DEFAULT_SHEET_VIEW_ID, TEMPORARY_SHEET_VIEW_ID } from '../../../commands/sheet-views.js';
import { dictionaries, type Strings } from '../../../i18n/strings.js';
import type { ExcelRibbonIconName } from '../../excel-ribbon-icons.js';
import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonOption,
  RibbonTabModel,
  ToolbarLang,
} from '../../ribbon-model.js';
import {
  EMBEDDED_UNSUPPORTED,
  type LabelLike,
  labels,
  macGroupTitle,
  macRibbonDisabledReason,
  macRibbonLabelForCommand,
  NOT_IMPLEMENTED,
  OFFICE_REQUIRED,
  t,
  text,
} from './labels.js';
import { MAC_RIBBON_MENU_COMMANDS, MAC_RIBBON_MENU_ITEMS, macMenuLeaves } from './menu-manifest.js';

export { macRibbonDisabledReason, macRibbonLabelForCommand } from './labels.js';
export {
  MAC_FORMULAS_MORE_MENU_ID,
  MAC_RIBBON_MENU_COMMAND_SET,
  MAC_RIBBON_MENU_COMMANDS,
  MAC_RIBBON_MENU_ITEMS,
  type MacRibbonMenuItem,
  macMenuLeaves,
  macRibbonMenuIdForCommand,
} from './menu-manifest.js';
export { projectMacRibbonState } from './state-projection.js';

const command = (
  id: string,
  label: LabelLike,
  icon?: string,
  options: {
    kind?: RibbonCommand['kind'];
    disabled?: boolean;
    disabledReason?: LabelLike;
    className?: string;
    layout?: RibbonCommand['layout'];
  } = {},
): RibbonCommand => ({
  id,
  title: text('en', label),
  label: text('en', label),
  icon,
  kind: options.kind ?? 'wide',
  layout: options.layout,
  className: options.className,
  disabled: options.disabled,
  disabledReason: options.disabledReason ? text('en', options.disabledReason) : undefined,
});

const group = (
  title: LabelLike,
  commands: RibbonCommand[],
  variant = 'tiles',
): RibbonGroupModel => ({
  title: text('en', title),
  commands,
  variant,
});

const menuCommand = (id: string, label: LabelLike, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'wide' });

const compactCommand = (id: string, label: LabelLike, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'wide', layout: 'stacked' });

const iconCommand = (id: string, label: LabelLike, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'button', className: 'fc-tb__rb--mac-icon-only' });

const disabledCompactCommand = (
  id: string,
  label: LabelLike,
  icon: string | undefined,
  reason: LabelLike,
): RibbonCommand =>
  command(id, label, icon, {
    kind: 'wide',
    layout: 'stacked',
    disabled: true,
    disabledReason: reason,
  });

const checkboxCommand = (id: string, label: LabelLike, icon?: string): RibbonCommand =>
  command(id, label, icon, {
    kind: 'wide',
    layout: 'stacked',
    className: 'fc-tb__rb--mac-checkbox',
  });

/** A checkbox the current view does not offer, as Ruler in Normal view. */
const disabledCheckboxCommand = (id: string, label: LabelLike, reason: LabelLike): RibbonCommand =>
  command(id, label, undefined, {
    kind: 'wide',
    layout: 'stacked',
    className: 'fc-tb__rb--mac-checkbox',
    disabled: true,
    disabledReason: reason,
  });

/** A static caption inside a group; it carries no action. */
const labelCommand = (id: string, label: LabelLike): RibbonCommand =>
  command(id, label, undefined, { kind: 'label' });

const selectCommand = (
  id: string,
  label: LabelLike,
  options: readonly RibbonOption[],
  className = 'fc-tb__rb-select--border',
): RibbonCommand => ({
  id,
  title: label.en,
  label: label.en,
  kind: 'select',
  options,
  className,
});

const disabledCommand = (
  id: string,
  label: LabelLike,
  icon: string | undefined,
  reason: LabelLike,
) => command(id, label, icon, { disabled: true, disabledReason: reason, kind: 'wide' });

const macTab = (
  id: RibbonTabModel['id'],
  label: LabelLike,
  groups: RibbonGroupModel[],
): RibbonTabModel => ({
  id,
  label: label.en,
  groups,
});
type FormulaFamily =
  | 'recent'
  | 'financial'
  | 'logical'
  | 'text'
  | 'dateTime'
  | 'lookup'
  | 'math'
  | 'more';

const FORMULA_FAMILY_KEYS = [
  'recent',
  'financial',
  'logical',
  'text',
  'dateTime',
  'lookup',
  'math',
  'more',
] as const satisfies readonly FormulaFamily[];

const FORMULA_FAMILY_ICONS = {
  recent: 'functionRecent',
  financial: 'functionFinancial',
  logical: 'functionLogical',
  text: 'functionText',
  dateTime: 'functionDateTime',
  lookup: 'functionLookup',
  math: 'functionMath',
  more: 'functionMore',
} as const satisfies Record<FormulaFamily, ExcelRibbonIconName>;

const insertTab = (): RibbonTabModel =>
  macTab('insert', t('insert'), [
    group({ ja: 'テーブル', en: 'Tables' }, [
      menuCommand('mac.insert.pivotTable', t('pivotTable'), 'pivotTable'),
      menuCommand('mac.insert.recommendedPivotTable', t('recommendedPivotTable'), 'pivotTable'),
      menuCommand('mac.insert.table', t('table'), 'table'),
    ]),
    group({ ja: 'フォーム', en: 'Forms' }, [
      disabledCommand('mac.insert.forms', t('forms'), 'form', OFFICE_REQUIRED),
    ]),
    group({ ja: '画像からのデータ', en: 'Data from Picture' }, [
      disabledCommand(
        'mac.insert.pictureFromData',
        t('pictureFromData'),
        'picture',
        OFFICE_REQUIRED,
      ),
    ]),
    group(
      { ja: 'イラスト', en: 'Illustrations' },
      [
        menuCommand('mac.insert.photo', t('photo'), 'picture'),
        menuCommand('mac.insert.shapes', t('shapes'), 'shapes'),
        disabledCommand('mac.insert.icons', t('icons'), 'icons', NOT_IMPLEMENTED),
        disabledCompactCommand(
          'mac.insert.threeDModel',
          t('threeDModel'),
          'threeD',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.insert.smartArt', t('smartArt'), 'smartArt', NOT_IMPLEMENTED),
        compactCommand('mac.insert.screenshot', t('screenshot'), 'screenshot'),
      ],
      'illustrations',
    ),
    group({ ja: 'コントロール', en: 'Controls' }, [
      disabledCommand('mac.insert.checkBox', t('checkBox'), 'checkBox', NOT_IMPLEMENTED),
    ]),
    group(
      { ja: 'グラフ', en: 'Charts' },
      [
        menuCommand('mac.insert.recommendedChart', t('recommendedChart'), 'chartRecommended'),
        iconCommand('mac.insert.chartColumn', t('chartColumn'), 'chartColumn'),
        iconCommand('mac.insert.chartLine', t('chartLine'), 'chartLine'),
        iconCommand('mac.insert.chartPie', t('chartPie'), 'chartPie'),
        iconCommand('mac.insert.chartHierarchy', t('chartHierarchy'), 'chartHierarchy'),
        iconCommand('mac.insert.chartStatistical', t('chartStatistical'), 'chartStatistical'),
        iconCommand('mac.insert.chartScatter', t('chartScatter'), 'chartScatter'),
        iconCommand('mac.insert.chartWaterfall', t('chartWaterfall'), 'chartWaterfall'),
        iconCommand('mac.insert.chartCombo', t('chartCombo'), 'chartCombo'),
        menuCommand('mac.insert.chartMap', t('chartMap'), 'chartMap'),
        disabledCommand('mac.insert.pivotChart', t('pivotChart'), 'pivotTable', NOT_IMPLEMENTED),
      ],
      'charts',
    ),
    group({ ja: 'スパークライン', en: 'Sparklines' }, [
      menuCommand('mac.insert.sparkline', t('sparkline'), 'sparkline'),
    ]),
    group({ ja: 'フィルター', en: 'Filters' }, [
      menuCommand('mac.insert.slicer', t('slicer'), 'slicer'),
      disabledCommand('mac.insert.timeline', t('timeline'), 'timeline', NOT_IMPLEMENTED),
    ]),
    group({ ja: 'リンク', en: 'Links' }, [menuCommand('mac.insert.link', t('link'), 'link')]),
    group({ ja: 'コメント', en: 'Comments' }, [
      command('mac.insert.newComment', t('newComment'), 'commentAdd'),
    ]),
    group({ ja: 'テキスト', en: 'Text' }, [
      disabledCommand('mac.insert.textBox', t('textBox'), 'textBox', EMBEDDED_UNSUPPORTED),
      menuCommand('mac.insert.headerFooter', t('headerFooter'), 'headerFooter'),
      disabledCommand('mac.insert.wordArt', t('wordArt'), 'wordArt', NOT_IMPLEMENTED),
      disabledCommand('mac.insert.object', t('object'), 'object', NOT_IMPLEMENTED),
    ]),
    group({ ja: '記号', en: 'Symbols' }, [
      disabledCommand('mac.insert.equation', t('equation'), 'function', NOT_IMPLEMENTED),
      menuCommand('mac.insert.symbol', t('symbol'), 'symbol'),
    ]),
  ]);

const drawTab = (): RibbonTabModel =>
  macTab('draw', t('draw'), [
    group(
      { ja: 'ツール', en: 'Tools' },
      [
        menuCommand('mac.draw.toggle', t('draw'), 'pen'),
        menuCommand('mac.draw.eraser', t('eraser'), 'eraser'),
        disabledCommand('mac.draw.lasso', t('lasso'), 'lasso', EMBEDDED_UNSUPPORTED),
      ],
      'draw-tools',
    ),
    group(
      { ja: 'ペン', en: 'Pens' },
      [
        iconCommand('mac.draw.penBlack', t('penBlack'), 'penBlack'),
        iconCommand('mac.draw.penRed', t('penRed'), 'penRed'),
        iconCommand('mac.draw.pencil', t('pencil'), 'pencil'),
        iconCommand('mac.draw.highlighter', t('highlighter'), 'highlighter'),
        menuCommand('mac.draw.add', t('add'), 'plus'),
      ],
      'pens',
    ),
    group(
      { ja: 'モード', en: 'Mode' },
      [menuCommand('mac.draw.trackpad', t('trackpad'), 'trackpad')],
      'draw-mode',
    ),
  ]);

const pageLayoutTab = (): RibbonTabModel =>
  macTab('pageLayout', { ja: 'ページ レイアウト', en: 'Page Layout' }, [
    group(
      { ja: 'テーマ', en: 'Themes' },
      [
        menuCommand('mac.page.theme', t('theme'), 'pageTheme'),
        compactCommand('mac.page.color', t('color'), 'pageTheme'),
        compactCommand('mac.page.font', t('font'), 'font'),
      ],
      'themes',
    ),
    group(
      { ja: 'ページ設定', en: 'Page Setup' },
      [
        menuCommand('mac.page.margins', t('margins'), 'pageSetup'),
        menuCommand('mac.page.orientation', t('orientation'), 'pageSetup'),
        menuCommand('mac.page.size', t('size'), 'pageSetup'),
        menuCommand('mac.page.printArea', t('printArea'), 'printArea'),
        menuCommand('mac.page.pageBreaks', t('pageBreaks'), 'pageBreaks'),
        menuCommand('mac.page.background', t('background'), 'sheetBackground'),
        menuCommand('mac.page.printTitles', t('printTitles'), 'printTitles'),
        menuCommand('mac.page.pageSetup', t('pageSetup'), 'pageSetup'),
      ],
      'page-setup',
    ),
    group({ ja: '拡大縮小印刷', en: 'Scale to Fit' }, [
      selectCommand('scaleWidth', t('scaleWidth'), [
        { value: '0', label: 'Automatic' },
        { value: '1', label: '1 page' },
        { value: '2', label: '2 pages' },
        { value: '3', label: '3 pages' },
        { value: 'custom', label: 'Custom' },
      ]),
      selectCommand('scaleHeight', t('scaleHeight'), [
        { value: '0', label: 'Automatic' },
        { value: '1', label: '1 page' },
        { value: '2', label: '2 pages' },
        { value: '3', label: '3 pages' },
        { value: 'custom', label: 'Custom' },
      ]),
    ]),
    group(
      { ja: 'シートのオプション', en: 'Sheet Options' },
      [
        labelCommand('mac.page.gridlinesHeading', t('gridlinesHeading')),
        checkboxCommand('mac.page.showGridlines', t('showGridlines')),
        checkboxCommand('mac.page.printGridlines', t('printGridlines')),
        labelCommand('mac.page.headingsHeading', t('headingsHeading')),
        checkboxCommand('mac.page.showHeadings', t('showHeadings')),
        checkboxCommand('mac.page.printHeadings', t('printHeadings')),
      ],
      'checks',
    ),
  ]);

const formulasTab = (): RibbonTabModel =>
  macTab('formulas', t('formulas'), [
    group(
      { ja: '関数ライブラリ', en: 'Function Library' },
      [
        menuCommand('mac.formulas.insertFunction', t('insertFunction'), 'function'),
        menuCommand('mac.formulas.autoSum', t('autoSum'), 'autosum'),
        ...FORMULA_FAMILY_KEYS.map((key) =>
          menuCommand(`mac.formulas.${key}`, t(key), FORMULA_FAMILY_ICONS[key]),
        ),
      ],
      'function-library',
    ),
    group(
      { ja: 'Python (プレビュー)', en: 'Python (Preview)' },
      [
        disabledCommand('mac.formulas.python', t('python'), 'python', OFFICE_REQUIRED),
        disabledCompactCommand('mac.formulas.reset', t('reset'), 'clear', OFFICE_REQUIRED),
        disabledCompactCommand('mac.formulas.editor', t('editor'), 'script', OFFICE_REQUIRED),
        disabledCompactCommand(
          'mac.formulas.initialize',
          t('initialize'),
          'script',
          OFFICE_REQUIRED,
        ),
      ],
      'python',
    ),
    group(
      { ja: '定義された名前', en: 'Defined Names' },
      [
        menuCommand('mac.formulas.namesManager', t('namesManager'), 'names'),
        compactCommand('mac.formulas.defineName', t('defineName'), 'names'),
        compactCommand('mac.formulas.useInFormula', t('useInFormula'), 'names'),
        compactCommand('mac.formulas.createFromSelection', t('createFromSelection'), 'names'),
      ],
      'defined-names',
    ),
    group(
      { ja: '数式の監査', en: 'Formula Auditing' },
      [
        compactCommand('mac.formulas.precedents', t('precedents'), 'trace'),
        compactCommand('mac.formulas.dependents', t('dependents'), 'dependents'),
        compactCommand('mac.formulas.removeArrows', t('removeArrows'), 'clearArrows'),
        menuCommand('mac.formulas.showFormulas', t('showFormulas'), 'function'),
        menuCommand('mac.formulas.errorCheck', t('errorCheck'), 'errorChecking'),
        menuCommand('mac.formulas.watch', t('watch'), 'watch'),
      ],
      'formula-auditing',
    ),
    group(
      { ja: '計算方法', en: 'Calculation' },
      [
        menuCommand('mac.formulas.calcOptions', t('calcOptions'), 'calcOptions'),
        compactCommand('mac.formulas.recalc', t('recalc'), 'autosum'),
        compactCommand('mac.formulas.sheetRecalc', t('sheetRecalc'), 'autosum'),
      ],
      'calculation',
    ),
  ]);

const dataTab = (): RibbonTabModel =>
  macTab('data', t('data'), [
    group({ ja: 'データの取得と変換', en: 'Get & Transform Data' }, [
      disabledCommand('mac.data.powerQuery', t('powerQuery'), 'data', NOT_IMPLEMENTED),
      menuCommand('mac.data.pictureFromData', t('pictureFromData'), 'picture'),
    ]),
    group(
      { ja: 'クエリと接続', en: 'Queries & Connections' },
      [
        disabledCompactCommand('mac.data.refreshAll', t('refreshAll'), 'refresh', NOT_IMPLEMENTED),
        disabledCompactCommand(
          'mac.data.queriesConnections',
          t('queriesConnections'),
          'link',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.data.properties', t('properties'), 'inspect', NOT_IMPLEMENTED),
        compactCommand('mac.data.workbookLinks', t('workbookLinks'), 'link'),
      ],
      'compact',
    ),
    group(
      { ja: 'データの種類', en: 'Data Types' },
      [menuCommand('mac.data.dataTypes', t('dataTypes'), 'data')],
      'compact',
    ),
    group(
      { ja: '並べ替えとフィルター', en: 'Sort & Filter' },
      [
        iconCommand('mac.data.sortAsc', t('sortAsc'), 'sortAsc'),
        iconCommand('mac.data.sortDesc', t('sortDesc'), 'sortDesc'),
        menuCommand('mac.data.sortCustom', t('sortCustom'), 'sortAsc'),
        menuCommand('mac.data.filter', t('filter'), 'filter'),
        compactCommand('mac.data.clear', t('clear'), 'clear'),
        compactCommand('mac.data.reapply', t('reapply'), 'refresh'),
        compactCommand('mac.data.advancedFilter', t('advancedFilter'), 'filter'),
      ],
      'sort-filter',
    ),
    group(
      { ja: 'データ ツール', en: 'Data Tools' },
      [
        compactCommand('mac.data.textToColumns', t('textToColumns'), 'textToColumns'),
        disabledCompactCommand('mac.data.dataCleanup', t('dataCleanup'), 'clear', OFFICE_REQUIRED),
        disabledCompactCommand(
          'mac.data.dataAnalysis',
          t('dataAnalysis'),
          'inspect',
          OFFICE_REQUIRED,
        ),
        compactCommand('mac.data.flashFill', t('flashFill'), 'flashFill'),
        compactCommand('mac.data.removeDuplicates', t('removeDuplicates'), 'removeDuplicates'),
        compactCommand('mac.data.validation', t('validation'), 'dataValidation'),
        compactCommand('mac.data.consolidate', t('consolidate'), 'merge'),
      ],
      'compact',
    ),
    group(
      { ja: '予測', en: 'Forecast' },
      [menuCommand('mac.data.whatIf', t('whatIf'), 'whatIf')],
      'compact',
    ),
    group(
      { ja: 'アウトライン', en: 'Outline' },
      [
        compactCommand('mac.data.group', t('group'), 'outlineGroup'),
        compactCommand('mac.data.ungroup', t('ungroup'), 'outlineUngroup'),
        compactCommand('mac.data.subtotal', t('subtotal'), 'outlineGroup'),
        compactCommand('mac.data.showDetail', t('showDetail'), 'outlineShow'),
        compactCommand('mac.data.hideDetail', t('hideDetail'), 'outlineHide'),
      ],
      'compact',
    ),
    group(
      { ja: '分析', en: 'Analysis' },
      [disabledCompactCommand('mac.data.analysis', t('analysis'), 'chart', NOT_IMPLEMENTED)],
      'compact',
    ),
  ]);

const reviewTab = (): RibbonTabModel =>
  macTab('review', t('review'), [
    group({ ja: '文章校正', en: 'Proofing' }, [
      menuCommand('mac.review.spelling', t('spelling'), 'spelling'),
      disabledCommand('mac.review.thesaurus', t('thesaurus'), 'spelling', EMBEDDED_UNSUPPORTED),
      menuCommand('mac.review.stats', t('stats'), 'count'),
    ]),
    group({ ja: 'アクセシビリティ', en: 'Accessibility' }, [
      menuCommand('mac.review.accessibility', t('accessibility'), 'accessibility'),
    ]),
    group({ ja: '言語', en: 'Language' }, [
      disabledCommand('mac.review.translate', t('translate'), 'translate', NOT_IMPLEMENTED),
    ]),
    group({ ja: '変更', en: 'Changes' }, [
      disabledCommand(
        'mac.review.showChanges',
        t('showChanges'),
        'showChanges',
        EMBEDDED_UNSUPPORTED,
      ),
    ]),
    group({ ja: 'コメント', en: 'Comments' }, [
      menuCommand('mac.review.newComment', t('newComment'), 'commentAdd'),
      menuCommand('mac.review.deleteComment', t('deleteComment'), 'clear'),
      menuCommand('mac.review.previousComment', t('previousComment'), 'goTo'),
      menuCommand('mac.review.nextComment', t('nextComment'), 'goTo'),
      menuCommand('mac.review.showComments', t('showComments'), 'commentAdd'),
    ]),
    group({ ja: 'メモ', en: 'Notes' }, [menuCommand('mac.review.notes', t('notes'), 'commentAdd')]),
    group({ ja: '変更箇所', en: 'Protection Changes' }, [
      menuCommand('mac.review.protectSheet', t('protectSheet'), 'protect'),
      menuCommand('mac.review.protectWorkbook', t('protectWorkbook'), 'protect'),
    ]),
    group({ ja: '保護', en: 'Protect' }, [
      disabledCommand('mac.review.alwaysReadOnly', t('alwaysReadOnly'), 'protect', NOT_IMPLEMENTED),
    ]),
    group({ ja: 'インク', en: 'Ink' }, [
      disabledCommand('mac.review.hideInk', t('hideInk'), 'pen', EMBEDDED_UNSUPPORTED),
    ]),
  ]);

const viewTab = (): RibbonTabModel =>
  macTab('view', t('view'), [
    group(
      { ja: 'シート ビュー', en: 'Sheet Views' },
      [
        selectCommand('sheetViewSelect', t('sheetViewSelect'), [
          { value: DEFAULT_SHEET_VIEW_ID, label: 'Default' },
          { value: TEMPORARY_SHEET_VIEW_ID, label: 'Temporary View' },
        ]),
        compactCommand('mac.view.sheetViewKeep', t('sheetViewKeep'), 'save'),
        compactCommand('mac.view.sheetViewExit', t('sheetViewExit'), 'paneClose'),
        compactCommand('mac.view.sheetViewNew', t('sheetViewNew'), 'plus'),
        compactCommand('mac.view.sheetViewOptions', t('sheetViewOptions'), 'options'),
      ],
      'compact',
    ),
    group(
      { ja: 'ブックの表示', en: 'Workbook Views' },
      [
        compactCommand('mac.view.standard', t('standard'), 'table'),
        compactCommand('mac.view.pageBreakPreview', t('pageBreakPreview'), 'table'),
        compactCommand('mac.view.pageLayout', t('pageLayout'), 'page'),
        disabledCompactCommand(
          'mac.view.customViews',
          t('customViews'),
          'customViews',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: '表示/非表示', en: 'Show' },
      [
        disabledCompactCommand('mac.view.navigation', t('navigation'), 'show', NOT_IMPLEMENTED),
        disabledCheckboxCommand('mac.view.ruler', t('ruler'), NOT_IMPLEMENTED),
        checkboxCommand('mac.view.gridlines', t('gridlines'), 'show'),
        checkboxCommand('mac.view.formulaBar', t('formulaBar'), 'show'),
        checkboxCommand('mac.view.headings', t('headings'), 'show'),
        disabledCompactCommand(
          'mac.view.dataTypeIcons',
          t('dataTypeIcons'),
          'show',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.view.cellFocus', t('cellFocus'), 'show', NOT_IMPLEMENTED),
      ],
      'compact',
    ),
    group(
      { ja: 'ズーム', en: 'Zoom' },
      [
        compactCommand('mac.view.zoom', t('zoom'), 'zoom'),
        compactCommand('mac.view.zoom100', t('zoom100'), 'zoom'),
        compactCommand('mac.view.zoomSelection', t('zoomSelection'), 'zoom'),
      ],
      'compact',
    ),
    group(
      { ja: 'ウィンドウ', en: 'Window' },
      [
        disabledCompactCommand(
          'mac.view.newWindow',
          t('newWindow'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand('mac.view.arrange', t('arrange'), 'arrange', EMBEDDED_UNSUPPORTED),
        compactCommand('mac.view.freeze', t('freeze'), 'freeze'),
        compactCommand('mac.view.split', t('split'), 'split'),
        disabledCompactCommand('mac.view.hide', t('hide'), 'hide', EMBEDDED_UNSUPPORTED),
        disabledCompactCommand('mac.view.unhide', t('unhide'), 'show', EMBEDDED_UNSUPPORTED),
        disabledCompactCommand(
          'mac.view.sideBySide',
          t('sideBySide'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.view.syncScroll',
          t('syncScroll'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.view.resetWindow',
          t('resetWindow'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.view.switchWindow',
          t('switchWindow'),
          'window',
          EMBEDDED_UNSUPPORTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'マクロ', en: 'Macros' },
      [
        disabledCompactCommand('mac.view.macros', t('macros'), 'script', NOT_IMPLEMENTED),
        disabledCompactCommand('mac.view.recordMacro', t('recordMacro'), 'script', NOT_IMPLEMENTED),
        disabledCompactCommand(
          'mac.view.relativeRecording',
          t('relativeRecording'),
          'script',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact',
    ),
  ]);

const automateTab = (): RibbonTabModel =>
  macTab('automate', t('automate'), [
    group({ ja: 'Office スクリプト', en: 'Office Scripts' }, [
      disabledCommand('mac.automate.newScript', t('newScript'), 'script', OFFICE_REQUIRED),
      menuCommand('mac.automate.showScripts', t('showScripts'), 'script'),
    ]),
    group(
      { ja: 'Office スクリプト ギャラリー', en: 'Office Scripts Gallery' },
      [
        compactCommand('mac.automate.allRowsColumns', t('allRowsColumns'), 'show'),
        compactCommand('mac.automate.freezeSelection', t('freezeSelection'), 'freeze'),
        compactCommand('mac.automate.makeSubtable', t('makeSubtable'), 'table'),
        compactCommand('mac.automate.removeHyperlinks', t('removeHyperlinks'), 'link'),
        compactCommand('mac.automate.countEmptyRows', t('countEmptyRows'), 'count'),
        compactCommand('mac.automate.tableToJson', t('tableToJson'), 'code'),
        compactCommand('mac.automate.newPivotTable', t('newPivotTable'), 'pivotTable'),
      ],
      'gallery',
    ),
  ]);

/** Icon-and-label rows: Clipboard's small commands and Alignment's wrap/merge. */
const MAC_HOME_WIDE_STACKED_COMMAND_IDS = new Set([
  'wrap',
  'merge',
  'cut',
  'copy',
  'formatPainter',
]);
/** Cells commands are full-height tiles on Mac rather than rows. */
const MAC_HOME_TILE_COMMAND_IDS = new Set(['insertRows', 'deleteRows', 'formatCellsHome']);
const MAC_HOME_LABEL_KEYS: Readonly<Record<string, keyof Strings['mac']['homeRibbon']>> = {
  paste: 'paste',
  cut: 'cut',
  copy: 'copy',
  formatPainter: 'formatPainter',
  conditional: 'conditional',
  cellStyles: 'cellStyles',
  autosum: 'autoSum',
};

/** Control names the Japanese Mac Home tab uses where they differ from the shared model. */
const MAC_HOME_JA_NAMES: Readonly<Record<string, string>> = {
  textOrientation: '文字の向き',
  indentDecrease: 'インデント解除',
  indentIncrease: 'インデント',
  borders: '線種/網かけ変更',
  fillColor: '網かけ',
  decUp: '小数点表示桁上げ',
  decDown: '小数点表示桁下げ',
  insertRows: '挿入',
  deleteRows: '削除',
};
/** Tooltips that read differently from the control name. */
const MAC_HOME_JA_HELP: Readonly<Record<string, string>> = {
  formatPainter: '書式を別の場所にコピーして適用',
  alignC: '文字列中央揃え',
};
/** Mac tooltips are the bare control name: no shortcut or symbol suffix. */
const stripShortcutHint = (title: string): string =>
  title.replace(/\s*\((?:Cmd|Ctrl|⌘)[^)]*\)|\s*\(Σ\)/g, '');

const macHomeCommand = (candidate: RibbonCommand, lang: ToolbarLang): RibbonCommand => {
  const measuredName = lang === 'ja' ? MAC_HOME_JA_NAMES[candidate.id] : undefined;
  const measuredHelp = lang === 'ja' ? MAC_HOME_JA_HELP[candidate.id] : undefined;
  const commandWithMacTitle = {
    ...candidate,
    ...(measuredName ? { label: measuredName, title: measuredName } : {}),
    title: stripShortcutHint(measuredName ?? candidate.title),
  };
  const macLabelKey = MAC_HOME_LABEL_KEYS[candidate.id];
  const homeLabel =
    candidate.id === 'merge'
      ? dictionaries[lang].ribbon.mergeAndCenter
      : candidate.id === 'wrap' && lang === 'ja'
        ? '折り返して全体を表示する'
        : macLabelKey
          ? dictionaries[lang].mac.homeRibbon[macLabelKey]
          : undefined;
  // Keep any title suffix such as AutoSum's "(Σ)" after the relabelled name.
  const macHomeTitle = (label: string): string =>
    commandWithMacTitle.title.startsWith(candidate.label)
      ? label + commandWithMacTitle.title.slice(candidate.label.length)
      : label;
  const commandWithMacHomeLabel = homeLabel
    ? { ...commandWithMacTitle, title: macHomeTitle(homeLabel), label: homeLabel }
    : commandWithMacTitle;
  if (measuredHelp) commandWithMacHomeLabel.title = measuredHelp;
  // The desktop font-size list starts at 6, and its percent row reads パーセンテージ.
  if (
    candidate.id === 'fontSize' &&
    candidate.options &&
    !candidate.options.some((o) => o.value === '6')
  ) {
    commandWithMacHomeLabel.options = [{ value: '6', label: '6' }, ...candidate.options];
  }
  if (candidate.id === 'numberFormat' && candidate.options && lang === 'ja') {
    commandWithMacHomeLabel.options = candidate.options.map((option) =>
      option.value === 'percent' ? { ...option, label: 'パーセンテージ' } : option,
    );
  }
  if (MAC_HOME_WIDE_STACKED_COMMAND_IDS.has(candidate.id)) {
    return {
      ...commandWithMacHomeLabel,
      kind: 'wide',
      layout: 'stacked',
    };
  }
  if (MAC_HOME_TILE_COMMAND_IDS.has(candidate.id)) {
    return { ...commandWithMacHomeLabel, kind: 'wide', layout: undefined };
  }
  return commandWithMacHomeLabel;
};

/** The shared Home tab's groups in order, keyed for their Mac group titles. */
const MAC_HOME_GROUP_KEYS = [
  'Clipboard',
  'Font',
  'Alignment',
  'Number',
  'Styles',
  'Cells',
  'Editing',
] as const;

const macHomeTab = (home: RibbonTabModel, lang: ToolbarLang): RibbonTabModel => ({
  ...home,
  groups: [
    ...home.groups.map((shared, index) => {
      const current = { ...shared, title: MAC_HOME_GROUP_KEYS[index] ?? shared.title };
      if (current.variant === 'cells') {
        // Tile layout rather than the shared stacked-rows layout.
        return {
          ...current,
          variant: 'tiles',
          commands: current.commands.map((candidate) => macHomeCommand(candidate, lang)),
        };
      }
      if (current.variant !== 'font') {
        return {
          ...current,
          commands: current.commands.map((candidate) => macHomeCommand(candidate, lang)),
        };
      }
      const isPhonetic = (id: string): boolean => id === 'showPhonetic';
      const phonetic = current.commands.filter((candidate) => isPhonetic(candidate.id));
      return {
        ...current,
        commands: [
          ...current.commands
            .filter((candidate) => candidate.id !== 'strike' && !isPhonetic(candidate.id))
            .map((candidate) => macHomeCommand(candidate, lang)),
          ...phonetic.map((candidate) => macHomeCommand(candidate, lang)),
        ],
      };
    }),
    group({ ja: 'アドイン', en: 'Add-ins' }, [
      menuCommand('mac.home.addins', t('addins'), 'addIn'),
    ]),
  ],
});

/** Builds the Office 365 Mac command surface. The shared Home model is passed
 *  in by ribbon-model.ts so its established formatting controls and active
 *  state projection remain identical across profiles. */
export const buildMacRibbonModel = (lang: ToolbarLang, home: RibbonTabModel): RibbonTabModel[] => {
  const localizeSelectOptions = (
    commandId: string,
    options: readonly RibbonOption[] | undefined,
  ): readonly RibbonOption[] | undefined => {
    if (!options) return undefined;
    if (
      commandId !== 'scaleWidth' &&
      commandId !== 'scaleHeight' &&
      commandId !== 'sheetViewSelect'
    ) {
      return options;
    }
    const strings = dictionaries[lang];
    const pageScale = strings.pageScale;
    const pageLabel = (value: string): string => {
      if (value === '0') return pageScale.automatic;
      if (value === 'custom') return pageScale.custom;
      return `${value} ${value === '1' ? pageScale.page : pageScale.pages}`;
    };
    return options.map((option) => ({
      ...option,
      label:
        commandId === 'sheetViewSelect'
          ? option.value === TEMPORARY_SHEET_VIEW_ID
            ? strings.mac.sheetViews.temporaryView
            : strings.mac.sheetViews.defaultView
          : pageLabel(option.value),
    }));
  };
  const localize = (tab: RibbonTabModel): RibbonTabModel => ({
    ...tab,
    label: labels[tab.id]?.[lang] ?? tab.label,
    groups: tab.groups.map((g) => ({
      ...g,
      title: macGroupTitle(g.title, lang),
      commands: g.commands.map((c) => ({
        ...c,
        title:
          macRibbonLabelForCommand(c.id, lang) === c.id
            ? c.title
            : macRibbonLabelForCommand(c.id, lang),
        label:
          macRibbonLabelForCommand(c.id, lang) === c.id
            ? c.label
            : macRibbonLabelForCommand(c.id, lang),
        disabledReason: macRibbonDisabledReason(c.disabledReason, lang),
        options: localizeSelectOptions(c.id, c.options),
      })),
    })),
  });
  return [
    macHomeTab(home, lang),
    insertTab(),
    drawTab(),
    pageLayoutTab(),
    formulasTab(),
    dataTab(),
    reviewTab(),
    viewTab(),
    automateTab(),
  ].map(localize);
};

const macCatalogHome: RibbonTabModel = { id: 'home', label: 'Home', groups: [] };
const macCatalog = buildMacRibbonModel('en', macCatalogHome).filter((tab) => tab.id !== 'home');
const macMenuLeafCommandIds = Object.values(MAC_RIBBON_MENU_ITEMS).flatMap((items) =>
  macMenuLeaves(items).map((item) => item.id),
);

export const MAC_RIBBON_COMMAND_IDS: readonly string[] = Object.freeze(
  Array.from(
    new Set([
      ...macCatalog.flatMap((tab) => tab.groups.flatMap((g) => g.commands.map((c) => c.id))),
      ...macMenuLeafCommandIds,
      ...MAC_RIBBON_MENU_COMMANDS,
      'mac.home.addins',
      ...allFunctionNames().map((name) => `mac.function.${name}`),
    ]),
  ),
);

export const MAC_RIBBON_DISABLED_COMMAND_IDS: ReadonlySet<string> = new Set([
  ...macCatalog.flatMap((tab) =>
    tab.groups.flatMap((g) => g.commands.filter((c) => c.disabled).map((c) => c.id)),
  ),
  ...Object.values(MAC_RIBBON_MENU_ITEMS).flatMap((items) =>
    macMenuLeaves(items)
      .filter((item) => item.disabled)
      .map((item) => item.id),
  ),
]);
