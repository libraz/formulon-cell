import { allFunctionNames } from '../../../commands/function-categories.js';
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
  GROUP_LABELS,
  type Label,
  labels,
  macRibbonDisabledReason,
  macRibbonLabelForCommand,
  NOT_IMPLEMENTED,
  OFFICE_REQUIRED,
  THREADED_COMMENTS_UNSUPPORTED,
  t,
  text,
  WORKBOOK_THEME_UNSUPPORTED,
} from './labels.js';
import { MAC_RIBBON_MENU_COMMANDS, MAC_RIBBON_MENU_ITEMS } from './menu-manifest.js';

export { macRibbonDisabledReason, macRibbonLabelForCommand } from './labels.js';
export {
  MAC_FORMULAS_MORE_MENU_ID,
  MAC_RIBBON_MENU_COMMAND_SET,
  MAC_RIBBON_MENU_COMMANDS,
  MAC_RIBBON_MENU_ITEMS,
  type MacRibbonMenuItem,
  macRibbonMenuIdForCommand,
} from './menu-manifest.js';
export { projectMacRibbonState } from './state-projection.js';

const command = (
  id: string,
  label: Label,
  icon?: string,
  options: {
    kind?: RibbonCommand['kind'];
    disabled?: boolean;
    disabledReason?: Label;
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

const group = (title: Label, commands: RibbonCommand[], variant = 'tiles'): RibbonGroupModel => ({
  title: text('en', title),
  commands,
  variant,
});

const menuCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'wide' });

const compactCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'wide', layout: 'stacked' });

const iconCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, { kind: 'button', className: 'fc-tb__rb--mac-icon-only' });

const disabledIconCommand = (
  id: string,
  label: Label,
  icon: string | undefined,
  reason: Label,
): RibbonCommand =>
  command(id, label, icon, {
    kind: 'button',
    className: 'fc-tb__rb--mac-icon-only',
    disabled: true,
    disabledReason: reason,
  });

const disabledCompactCommand = (
  id: string,
  label: Label,
  icon: string | undefined,
  reason: Label,
): RibbonCommand =>
  command(id, label, icon, {
    kind: 'wide',
    layout: 'stacked',
    disabled: true,
    disabledReason: reason,
  });

const checkboxCommand = (id: string, label: Label, icon?: string): RibbonCommand =>
  command(id, label, icon, {
    kind: 'wide',
    layout: 'stacked',
    className: 'fc-tb__rb--mac-checkbox',
  });

const selectCommand = (
  id: string,
  label: Label,
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

const disabledCommand = (id: string, label: Label, icon: string | undefined, reason: Label) =>
  command(id, label, icon, { disabled: true, disabledReason: reason, kind: 'wide' });

const macTab = (
  id: RibbonTabModel['id'],
  label: Label,
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
    group(
      { ja: '画像', en: 'Pictures' },
      [
        disabledCompactCommand(
          'mac.insert.pictureFromData',
          t('pictureFromData'),
          'picture',
          OFFICE_REQUIRED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'イラスト', en: 'Illustrations' },
      [
        compactCommand('mac.insert.photo', t('photo'), 'picture'),
        compactCommand('mac.insert.shapes', t('shapes'), 'shapes'),
        disabledCompactCommand('mac.insert.icons', t('icons'), 'icons', NOT_IMPLEMENTED),
        disabledCompactCommand(
          'mac.insert.threeDModel',
          t('threeDModel'),
          'threeD',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.insert.smartArt', t('smartArt'), 'smartArt', NOT_IMPLEMENTED),
        compactCommand('mac.insert.screenshot', t('screenshot'), 'screenshot'),
      ],
      'compact',
    ),
    group({ ja: 'コントロール', en: 'Controls' }, [
      disabledCommand('mac.insert.checkBox', t('checkBox'), 'checkBox', NOT_IMPLEMENTED),
    ]),
    group(
      { ja: 'グラフ', en: 'Charts' },
      [
        iconCommand('mac.insert.recommendedChart', t('recommendedChart'), 'chartRecommended'),
        iconCommand('mac.insert.chartColumn', t('chartColumn'), 'chartColumn'),
        iconCommand('mac.insert.chartLine', t('chartLine'), 'chartLine'),
        iconCommand('mac.insert.chartPie', t('chartPie'), 'chartPie'),
        disabledIconCommand(
          'mac.insert.chartHierarchy',
          t('chartHierarchy'),
          'chartHierarchy',
          NOT_IMPLEMENTED,
        ),
        disabledIconCommand(
          'mac.insert.chartStatistical',
          t('chartStatistical'),
          'chartStatistical',
          NOT_IMPLEMENTED,
        ),
        iconCommand('mac.insert.chartScatter', t('chartScatter'), 'chartScatter'),
        disabledIconCommand(
          'mac.insert.chartWaterfall',
          t('chartWaterfall'),
          'chartWaterfall',
          NOT_IMPLEMENTED,
        ),
        disabledIconCommand(
          'mac.insert.chartCombo',
          t('chartCombo'),
          'chartCombo',
          NOT_IMPLEMENTED,
        ),
        disabledIconCommand('mac.insert.chartMap', t('chartMap'), 'chartMap', OFFICE_REQUIRED),
        disabledIconCommand(
          'mac.insert.pivotChart',
          t('pivotChart'),
          'pivotTable',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact-icons',
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
      disabledCommand(
        'mac.insert.newComment',
        t('newComment'),
        'commentAdd',
        THREADED_COMMENTS_UNSUPPORTED,
      ),
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
        compactCommand('mac.draw.toggle', t('draw'), 'pen'),
        compactCommand('mac.draw.eraser', t('eraser'), 'eraser'),
        disabledCompactCommand('mac.draw.lasso', t('lasso'), 'lasso', EMBEDDED_UNSUPPORTED),
      ],
      'compact',
    ),
    group(
      { ja: 'ペン', en: 'Pens' },
      [
        iconCommand('mac.draw.penBlack', t('penBlack'), 'penBlack'),
        iconCommand('mac.draw.penRed', t('penRed'), 'penRed'),
        iconCommand('mac.draw.pencil', t('pencil'), 'pencil'),
        iconCommand('mac.draw.highlighter', t('highlighter'), 'highlighter'),
        iconCommand('mac.draw.add', t('add'), 'plus'),
      ],
      'pens',
    ),
    group(
      { ja: '描画', en: 'Draw' },
      [compactCommand('mac.draw.trackpad', t('trackpad'), 'trackpad')],
      'compact',
    ),
  ]);

const pageLayoutTab = (): RibbonTabModel =>
  macTab('pageLayout', { ja: 'ページ レイアウト', en: 'Page Layout' }, [
    group(
      { ja: 'テーマ', en: 'Themes' },
      [
        disabledCommand('mac.page.theme', t('theme'), 'pageTheme', WORKBOOK_THEME_UNSUPPORTED),
        disabledCommand('mac.page.color', t('color'), 'pageTheme', WORKBOOK_THEME_UNSUPPORTED),
        disabledCommand('mac.page.font', t('font'), 'font', WORKBOOK_THEME_UNSUPPORTED),
      ],
      'compact',
    ),
    group(
      { ja: 'ページ設定', en: 'Page Setup' },
      [
        compactCommand('mac.page.margins', t('margins'), 'pageSetup'),
        compactCommand('mac.page.orientation', t('orientation'), 'pageSetup'),
        compactCommand('mac.page.size', t('size'), 'pageSetup'),
        compactCommand('mac.page.printArea', t('printArea'), 'printArea'),
        compactCommand('mac.page.pageBreaks', t('pageBreaks'), 'pageBreaks'),
        compactCommand('mac.page.background', t('background'), 'sheetBackground'),
        compactCommand('mac.page.printTitles', t('printTitles'), 'printTitles'),
        compactCommand('mac.page.pageSetup', t('pageSetup'), 'pageSetup'),
      ],
      'compact',
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
        checkboxCommand('mac.page.showGridlines', t('showGridlines')),
        checkboxCommand('mac.page.printGridlines', t('printGridlines')),
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
      { ja: 'Python', en: 'Python' },
      [
        disabledCommand('mac.formulas.python', t('python'), 'python', OFFICE_REQUIRED),
        disabledCommand('mac.formulas.reset', t('reset'), 'clear', OFFICE_REQUIRED),
        disabledCommand('mac.formulas.editor', t('editor'), 'script', OFFICE_REQUIRED),
        disabledCommand('mac.formulas.initialize', t('initialize'), 'script', OFFICE_REQUIRED),
      ],
      'compact',
    ),
    group(
      { ja: '定義された名前', en: 'Defined Names' },
      [
        menuCommand('mac.formulas.namesManager', t('namesManager'), 'names'),
        menuCommand('mac.formulas.defineName', t('defineName'), 'names'),
        menuCommand('mac.formulas.useInFormula', t('useInFormula'), 'names'),
        menuCommand('mac.formulas.createFromSelection', t('createFromSelection'), 'names'),
      ],
      'compact',
    ),
    group(
      { ja: '数式の監査', en: 'Formula Auditing' },
      [
        menuCommand('mac.formulas.precedents', t('precedents'), 'trace'),
        menuCommand('mac.formulas.dependents', t('dependents'), 'dependents'),
        menuCommand('mac.formulas.removeArrows', t('removeArrows'), 'clearArrows'),
        menuCommand('mac.formulas.showFormulas', t('showFormulas'), 'function'),
        menuCommand('mac.formulas.errorCheck', t('errorCheck'), 'errorChecking'),
        menuCommand('mac.formulas.watch', t('watch'), 'watch'),
      ],
      'compact',
    ),
    group(
      { ja: '計算方法', en: 'Calculation' },
      [
        menuCommand('mac.formulas.calcOptions', t('calcOptions'), 'calcOptions'),
        menuCommand('mac.formulas.recalc', t('recalc'), 'autosum'),
        menuCommand('mac.formulas.sheetRecalc', t('sheetRecalc'), 'autosum'),
      ],
      'compact',
    ),
  ]);

const dataTab = (): RibbonTabModel =>
  macTab('data', t('data'), [
    group({ ja: 'データの取得と変換', en: 'Get & Transform Data' }, [
      disabledCommand('mac.data.powerQuery', t('powerQuery'), 'data', NOT_IMPLEMENTED),
      disabledCommand('mac.data.pictureFromData', t('pictureFromData'), 'picture', OFFICE_REQUIRED),
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
        disabledCompactCommand('mac.data.properties', t('properties'), 'options', NOT_IMPLEMENTED),
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
        compactCommand('mac.data.sortAsc', t('sortAsc'), 'sortAsc'),
        compactCommand('mac.data.sortDesc', t('sortDesc'), 'sortDesc'),
        compactCommand('mac.data.sortCustom', t('sortCustom'), 'sortAsc'),
        checkboxCommand('mac.data.filter', t('filter'), 'filter'),
        compactCommand('mac.data.clear', t('clear'), 'clear'),
        compactCommand('mac.data.reapply', t('reapply'), 'refresh'),
        compactCommand('mac.data.advancedFilter', t('advancedFilter'), 'filter'),
      ],
      'compact',
    ),
    group(
      { ja: 'データ ツール', en: 'Data Tools' },
      [
        compactCommand('mac.data.textToColumns', t('textToColumns'), 'textToColumns'),
        compactCommand('mac.data.flashFill', t('flashFill'), 'flashFill'),
        compactCommand('mac.data.removeDuplicates', t('removeDuplicates'), 'removeDuplicates'),
        compactCommand('mac.data.validation', t('validation'), 'dataValidation'),
        compactCommand('mac.data.consolidate', t('consolidate'), 'merge'),
      ],
      'compact',
    ),
    group(
      { ja: '予測', en: 'Forecast' },
      [menuCommand('mac.data.whatIf', t('whatIf'), 'options')],
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
    group(
      { ja: '文章校正', en: 'Proofing' },
      [
        compactCommand('mac.review.spelling', t('spelling'), 'spelling'),
        disabledCompactCommand(
          'mac.review.thesaurus',
          t('thesaurus'),
          'spelling',
          EMBEDDED_UNSUPPORTED,
        ),
        compactCommand('mac.review.stats', t('stats'), 'options'),
      ],
      'compact',
    ),
    group(
      { ja: 'アクセシビリティ', en: 'Accessibility' },
      [compactCommand('mac.review.accessibility', t('accessibility'), 'accessibility')],
      'compact',
    ),
    group(
      { ja: '言語', en: 'Language' },
      [
        disabledCompactCommand(
          'mac.review.translate',
          t('translate'),
          'translate',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: '変更', en: 'Changes' },
      [
        disabledCompactCommand(
          'mac.review.showChanges',
          t('showChanges'),
          'options',
          EMBEDDED_UNSUPPORTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'コメント', en: 'Comments' },
      [
        disabledCompactCommand(
          'mac.review.newComment',
          t('newComment'),
          'commentAdd',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.deleteComment',
          t('deleteComment'),
          'clear',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.previousComment',
          t('previousComment'),
          'goTo',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.nextComment',
          t('nextComment'),
          'goTo',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
        disabledCompactCommand(
          'mac.review.showComments',
          t('showComments'),
          'commentAdd',
          THREADED_COMMENTS_UNSUPPORTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: 'メモ', en: 'Notes' },
      [compactCommand('mac.review.notes', t('notes'), 'commentAdd')],
      'compact',
    ),
    group(
      { ja: '保護', en: 'Protect' },
      [
        compactCommand('mac.review.protectSheet', t('protectSheet'), 'protect'),
        compactCommand('mac.review.protectWorkbook', t('protectWorkbook'), 'protect'),
      ],
      'compact',
    ),
    group(
      { ja: 'オプション', en: 'Options' },
      [
        disabledCompactCommand(
          'mac.review.alwaysReadOnly',
          t('alwaysReadOnly'),
          'protect',
          NOT_IMPLEMENTED,
        ),
        disabledCompactCommand('mac.review.hideInk', t('hideInk'), 'pen', EMBEDDED_UNSUPPORTED),
      ],
      'compact',
    ),
  ]);

const viewTab = (): RibbonTabModel =>
  macTab('view', t('view'), [
    group(
      { ja: 'シート ビュー', en: 'Sheet Views' },
      [
        selectCommand('sheetViewSelect', t('sheetViewSelect'), [
          { value: 'current', label: 'Current View' },
        ]),
        compactCommand('mac.view.sheetViewSave', t('sheetViewSave'), 'options'),
        compactCommand('mac.view.sheetViewDelete', t('sheetViewDelete'), 'clear'),
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
          'options',
          NOT_IMPLEMENTED,
        ),
      ],
      'compact',
    ),
    group(
      { ja: '表示/非表示', en: 'Show' },
      [compactCommand('mac.view.show', t('show'), 'options')],
      'compact',
    ),
    group(
      { ja: 'ズーム', en: 'Zoom' },
      [
        compactCommand('mac.view.zoom', t('zoom'), 'zoom'),
        compactCommand('mac.view.zoom100', { ja: '100%', en: '100%' }, 'zoom'),
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
        disabledCompactCommand('mac.view.split', t('split'), 'split', EMBEDDED_UNSUPPORTED),
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
      [disabledCompactCommand('mac.view.macros', t('macros'), 'script', NOT_IMPLEMENTED)],
      'compact',
    ),
  ]);

const automateTab = (): RibbonTabModel =>
  macTab('automate', t('automate'), [
    group(
      { ja: '自動化', en: 'Automation' },
      [
        disabledCompactCommand('mac.automate.newScript', t('newScript'), 'script', OFFICE_REQUIRED),
        compactCommand('mac.automate.showScripts', t('showScripts'), 'script'),
        menuCommand('mac.automate.gallery', t('gallery'), 'script'),
      ],
      'compact',
    ),
    group(
      { ja: 'ギャラリー', en: 'Gallery' },
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

const MAC_HOME_ICON_ONLY_COMMAND_IDS = new Set(['autosum', 'fillHome', 'clearFormat']);

const macHomeCommand = (candidate: RibbonCommand): RibbonCommand => {
  const commandWithMacTitle = {
    ...candidate,
    title: candidate.title.replace(/\bCtrl\+/g, 'Cmd+'),
  };
  if (!MAC_HOME_ICON_ONLY_COMMAND_IDS.has(candidate.id)) return commandWithMacTitle;
  return {
    ...commandWithMacTitle,
    kind: 'button',
    layout: 'stacked',
    className: [candidate.className, 'fc-tb__rb--mac-icon-only'].filter(Boolean).join(' '),
  };
};

const macHomeTab = (home: RibbonTabModel): RibbonTabModel => ({
  ...home,
  groups: [
    ...home.groups.map((current) => {
      if (current.variant !== 'font') {
        return {
          ...current,
          commands: current.commands.map(macHomeCommand),
        };
      }
      const phonetic = current.commands.find((candidate) => candidate.id === 'editPhonetic');
      return {
        ...current,
        commands: [
          ...current.commands
            .filter((candidate) => candidate.id !== 'strike' && candidate.id !== 'editPhonetic')
            .map(macHomeCommand),
          ...(phonetic ? [phonetic] : []),
        ],
      };
    }),
    group({ ja: 'アドイン', en: 'Add-ins' }, [
      disabledCommand('mac.home.addins', t('addins'), 'addIn', OFFICE_REQUIRED),
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
    const pageLabel = (value: string): string => {
      if (lang === 'en') {
        if (value === '0') return 'Automatic';
        if (value === 'custom') return 'Custom';
        return `${value} ${value === '1' ? 'page' : 'pages'}`;
      }
      if (value === '0') return '自動';
      if (value === 'custom') return 'ユーザー設定';
      return `${value} ページ`;
    };
    return options.map((option) => ({
      ...option,
      label:
        commandId === 'sheetViewSelect' && option.value === 'current'
          ? lang === 'ja'
            ? '現在のビュー'
            : 'Current View'
          : pageLabel(option.value),
    }));
  };
  const localize = (tab: RibbonTabModel): RibbonTabModel => ({
    ...tab,
    label: labels[tab.id]?.[lang] ?? tab.label,
    groups: tab.groups.map((g) => ({
      ...g,
      title: GROUP_LABELS[g.title]?.[lang] ?? g.title,
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
    macHomeTab(home),
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
  items.map((item) => item.id),
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
  'mac.home.addins',
  ...macCatalog.flatMap((tab) =>
    tab.groups.flatMap((g) => g.commands.filter((c) => c.disabled).map((c) => c.id)),
  ),
  ...Object.values(MAC_RIBBON_MENU_ITEMS).flatMap((items) =>
    items.filter((item) => item.disabled).map((item) => item.id),
  ),
]);
