import { macFunctionCommandId } from '../../../commands/built-in-command-policy.js';
import { allFunctionNames } from '../../../commands/function-categories.js';
import { DEFAULT_SHEET_VIEW_ID, TEMPORARY_SHEET_VIEW_ID } from '../../../commands/sheet-views.js';
import { dictionaries, type Strings } from '../../../i18n/strings.js';
import type { ExcelRibbonIconName } from '../../excel-ribbon-icons.js';
import type {
  RibbonCommand,
  RibbonGroupModel,
  RibbonTabModel,
  ToolbarLang,
} from '../../ribbon-model.js';
import { findMacBundleLabel } from './bundle-text.js';
import { labels, macGroupTitle, macRibbonLabelForCommand, t } from './labels.js';
import { MAC_CONTEXT_LABELS, type MacRibbonReason } from './labels-data.js';
import { MAC_RIBBON_MENU_COMMANDS, MAC_RIBBON_MENU_ITEMS, macMenuLeaves } from './menu-manifest.js';

export { macRibbonLabelForCommand } from './labels.js';
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

/** Select options by value; their text comes from the dictionaries when the model is localized. */
interface MacSelectOptions {
  readonly values: readonly string[];
  readonly text: (value: string, strings: Strings) => string;
}

/** A Mac command before localization: its title and label come from its id, its reason from a
 *  dictionary key. */
type MacCommandDraft = Omit<RibbonCommand, 'title' | 'label' | 'disabledReason' | 'options'> & {
  reason?: MacRibbonReason;
  options?: MacSelectOptions;
};

interface MacGroupDraft {
  title: string;
  commands: MacCommandDraft[];
  variant: string;
}

interface MacTabDraft {
  id: RibbonTabModel['id'];
  groups: MacGroupDraft[];
}

const PAGE_SCALE_OPTIONS: MacSelectOptions = {
  values: ['0', '1', '2', '3', 'custom'],
  text: (value, { pageScale }) => {
    if (value === '0') return pageScale.automatic;
    if (value === 'custom') return pageScale.custom;
    return `${value} ${value === '1' ? pageScale.page : pageScale.pages}`;
  },
};

const SHEET_VIEW_OPTIONS: MacSelectOptions = {
  values: [DEFAULT_SHEET_VIEW_ID, TEMPORARY_SHEET_VIEW_ID],
  text: (value, { mac }) =>
    value === TEMPORARY_SHEET_VIEW_ID ? mac.sheetViews.temporaryView : mac.sheetViews.defaultView,
};

const command = (
  id: string,
  icon?: string,
  options: {
    kind?: RibbonCommand['kind'];
    disabled?: boolean;
    reason?: MacRibbonReason;
    className?: string;
    layout?: RibbonCommand['layout'];
  } = {},
): MacCommandDraft => ({
  id,
  icon,
  kind: options.kind ?? 'wide',
  layout: options.layout,
  className: options.className,
  disabled: options.disabled,
  reason: options.reason,
});

/** A group titled by its `groups:` label key, which the localize pass turns into the title. */
const group = (key: string, commands: MacCommandDraft[], variant = 'tiles'): MacGroupDraft => ({
  title: key,
  commands,
  variant,
});

const menuCommand = (id: string, icon?: string): MacCommandDraft =>
  command(id, icon, { kind: 'wide' });

const compactCommand = (id: string, icon?: string): MacCommandDraft =>
  command(id, icon, { kind: 'wide', layout: 'stacked' });

const iconCommand = (id: string, icon?: string): MacCommandDraft =>
  command(id, icon, { kind: 'button', className: 'fc-tb__rb--mac-icon-only' });

const disabledCompactCommand = (
  id: string,
  icon: string | undefined,
  reason: MacRibbonReason,
): MacCommandDraft =>
  command(id, icon, { kind: 'wide', layout: 'stacked', disabled: true, reason });

const checkboxCommand = (id: string, icon?: string): MacCommandDraft =>
  command(id, icon, { kind: 'wide', layout: 'stacked', className: 'fc-tb__rb--mac-checkbox' });

/** A checkbox the current view does not offer, as Ruler in Normal view. */
const disabledCheckboxCommand = (id: string, reason: MacRibbonReason): MacCommandDraft =>
  command(id, undefined, {
    kind: 'wide',
    layout: 'stacked',
    className: 'fc-tb__rb--mac-checkbox',
    disabled: true,
    reason,
  });

/** A static caption inside a group; it carries no action. */
const labelCommand = (id: string): MacCommandDraft => command(id, undefined, { kind: 'label' });

const selectCommand = (
  id: string,
  options: MacSelectOptions,
  className = 'fc-tb__rb-select--border',
): MacCommandDraft => ({ id, kind: 'select', options, className });

const disabledCommand = (id: string, icon: string | undefined, reason: MacRibbonReason) =>
  command(id, icon, { disabled: true, reason, kind: 'wide' });

const macTab = (id: RibbonTabModel['id'], groups: MacGroupDraft[]): MacTabDraft => ({
  id,
  groups,
});

/** The text of a command: its id's label in `lang`. */
const localizeCommand = (draft: MacCommandDraft, lang: ToolbarLang): RibbonCommand => {
  const { id, reason, ...rest } = draft;
  const label = macRibbonLabelForCommand(id, lang);
  const options = draft.options;
  return {
    id,
    title: label,
    label,
    ...rest,
    disabledReason: reason && MAC_CONTEXT_LABELS[reason][lang],
    options: options?.values.map((value) => ({
      value,
      label: options.text(value, dictionaries[lang]),
    })),
  };
};

const localizeGroup = (draft: MacGroupDraft, lang: ToolbarLang): RibbonGroupModel => ({
  ...draft,
  title: macGroupTitle(draft.title, lang),
  commands: draft.commands.map((c) => localizeCommand(c, lang)),
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

const insertTab = (): MacTabDraft =>
  macTab('insert', [
    group('Tables', [
      menuCommand('mac.insert.pivotTable', 'pivotTable'),
      menuCommand('mac.insert.recommendedPivotTable', 'pivotTable'),
      menuCommand('mac.insert.table', 'table'),
    ]),
    group('Forms', [disabledCommand('mac.insert.forms', 'form', 'officeRequired')]),
    group('Data from Picture', [
      disabledCommand('mac.insert.pictureFromData', 'picture', 'officeRequired'),
    ]),
    group(
      'Illustrations',
      [
        menuCommand('mac.insert.photo', 'picture'),
        menuCommand('mac.insert.shapes', 'shapes'),
        disabledCommand('mac.insert.icons', 'icons', 'notImplemented'),
        disabledCompactCommand('mac.insert.threeDModel', 'threeD', 'notImplemented'),
        disabledCompactCommand('mac.insert.smartArt', 'smartArt', 'notImplemented'),
        compactCommand('mac.insert.screenshot', 'screenshot'),
      ],
      'illustrations',
    ),
    group('Controls', [disabledCommand('mac.insert.checkBox', 'checkBox', 'notImplemented')]),
    group(
      'Charts',
      [
        menuCommand('mac.insert.recommendedChart', 'chartRecommended'),
        iconCommand('mac.insert.chartColumn', 'chartColumn'),
        iconCommand('mac.insert.chartLine', 'chartLine'),
        iconCommand('mac.insert.chartPie', 'chartPie'),
        iconCommand('mac.insert.chartHierarchy', 'chartHierarchy'),
        iconCommand('mac.insert.chartStatistical', 'chartStatistical'),
        iconCommand('mac.insert.chartScatter', 'chartScatter'),
        iconCommand('mac.insert.chartWaterfall', 'chartWaterfall'),
        iconCommand('mac.insert.chartCombo', 'chartCombo'),
        menuCommand('mac.insert.chartMap', 'chartMap'),
        disabledCommand('mac.insert.pivotChart', 'pivotTable', 'notImplemented'),
      ],
      'charts',
    ),
    group('Sparklines', [menuCommand('mac.insert.sparkline', 'sparkline')]),
    group('Filters', [
      menuCommand('mac.insert.slicer', 'slicer'),
      disabledCommand('mac.insert.timeline', 'timeline', 'notImplemented'),
    ]),
    group('Links', [menuCommand('mac.insert.link', 'link')]),
    group('Comments', [command('mac.insert.newComment', 'commentAdd')]),
    group('Text', [
      menuCommand('mac.insert.textBox', 'textBox'),
      menuCommand('mac.insert.headerFooter', 'headerFooter'),
      disabledCommand('mac.insert.wordArt', 'wordArt', 'notImplemented'),
      disabledCommand('mac.insert.object', 'object', 'notImplemented'),
    ]),
    group('Symbols', [
      disabledCommand('mac.insert.equation', 'function', 'notImplemented'),
      menuCommand('mac.insert.symbol', 'symbol'),
    ]),
  ]);

const drawTab = (): MacTabDraft =>
  macTab('draw', [
    group(
      'Tools',
      [
        menuCommand('mac.draw.toggle', 'pen'),
        menuCommand('mac.draw.eraser', 'eraser'),
        disabledCommand('mac.draw.lasso', 'lasso', 'embeddedUnsupported'),
      ],
      'draw-tools',
    ),
    group(
      'Pens',
      [
        iconCommand('mac.draw.penBlack', 'penBlack'),
        iconCommand('mac.draw.penRed', 'penRed'),
        iconCommand('mac.draw.pencil', 'pencil'),
        iconCommand('mac.draw.highlighter', 'highlighter'),
        menuCommand('mac.draw.add', 'plus'),
      ],
      'pens',
    ),
    group('Mode', [menuCommand('mac.draw.trackpad', 'trackpad')], 'draw-mode'),
  ]);

const pageLayoutTab = (): MacTabDraft =>
  macTab('pageLayout', [
    group(
      'Themes',
      [
        menuCommand('mac.page.theme', 'pageTheme'),
        compactCommand('mac.page.color', 'pageTheme'),
        compactCommand('mac.page.font', 'font'),
      ],
      'themes',
    ),
    group(
      'Page Setup',
      [
        menuCommand('mac.page.margins', 'pageSetup'),
        menuCommand('mac.page.orientation', 'pageSetup'),
        menuCommand('mac.page.size', 'pageSetup'),
        menuCommand('mac.page.printArea', 'printArea'),
        menuCommand('mac.page.pageBreaks', 'pageBreaks'),
        menuCommand('mac.page.background', 'sheetBackground'),
        menuCommand('mac.page.printTitles', 'printTitles'),
        menuCommand('mac.page.pageSetup', 'pageSetup'),
      ],
      'page-setup',
    ),
    group('Scale to Fit', [
      selectCommand('scaleWidth', PAGE_SCALE_OPTIONS),
      selectCommand('scaleHeight', PAGE_SCALE_OPTIONS),
    ]),
    group(
      'Sheet Options',
      [
        labelCommand('mac.page.gridlinesHeading'),
        checkboxCommand('mac.page.showGridlines'),
        checkboxCommand('mac.page.printGridlines'),
        labelCommand('mac.page.headingsHeading'),
        checkboxCommand('mac.page.showHeadings'),
        checkboxCommand('mac.page.printHeadings'),
      ],
      'checks',
    ),
  ]);

const formulasTab = (): MacTabDraft =>
  macTab('formulas', [
    group(
      'Function Library',
      [
        menuCommand('mac.formulas.insertFunction', 'function'),
        menuCommand('mac.formulas.autoSum', 'autosum'),
        ...FORMULA_FAMILY_KEYS.map((key) =>
          menuCommand(`mac.formulas.${key}`, FORMULA_FAMILY_ICONS[key]),
        ),
      ],
      'function-library',
    ),
    group(
      'Python (Preview)',
      [
        disabledCommand('mac.formulas.python', 'python', 'officeRequired'),
        disabledCompactCommand('mac.formulas.reset', 'clear', 'officeRequired'),
        disabledCompactCommand('mac.formulas.editor', 'script', 'officeRequired'),
        disabledCompactCommand('mac.formulas.initialize', 'script', 'officeRequired'),
      ],
      'python',
    ),
    group(
      'Defined Names',
      [
        menuCommand('mac.formulas.namesManager', 'names'),
        compactCommand('mac.formulas.defineName', 'names'),
        compactCommand('mac.formulas.useInFormula', 'names'),
        compactCommand('mac.formulas.createFromSelection', 'names'),
      ],
      'defined-names',
    ),
    group(
      'Formula Auditing',
      [
        compactCommand('mac.formulas.precedents', 'trace'),
        compactCommand('mac.formulas.dependents', 'dependents'),
        compactCommand('mac.formulas.removeArrows', 'clearArrows'),
        menuCommand('mac.formulas.showFormulas', 'function'),
        menuCommand('mac.formulas.errorCheck', 'errorChecking'),
        menuCommand('mac.formulas.watch', 'watch'),
      ],
      'formula-auditing',
    ),
    group(
      'Calculation',
      [
        menuCommand('mac.formulas.calcOptions', 'calcOptions'),
        compactCommand('mac.formulas.recalc', 'autosum'),
        compactCommand('mac.formulas.sheetRecalc', 'autosum'),
      ],
      'calculation',
    ),
  ]);

const dataTab = (): MacTabDraft =>
  macTab('data', [
    group('Get & Transform Data', [
      disabledCommand('mac.data.powerQuery', 'data', 'notImplemented'),
      menuCommand('mac.data.pictureFromData', 'picture'),
    ]),
    group(
      'Queries & Connections',
      [
        disabledCompactCommand('mac.data.refreshAll', 'refresh', 'notImplemented'),
        disabledCompactCommand('mac.data.queriesConnections', 'link', 'notImplemented'),
        disabledCompactCommand('mac.data.properties', 'inspect', 'notImplemented'),
        compactCommand('mac.data.workbookLinks', 'link'),
      ],
      'compact',
    ),
    group('Data Types', [menuCommand('mac.data.dataTypes', 'data')], 'compact'),
    group(
      'Sort & Filter',
      [
        iconCommand('mac.data.sortAsc', 'sortAsc'),
        iconCommand('mac.data.sortDesc', 'sortDesc'),
        menuCommand('mac.data.sortCustom', 'sortAsc'),
        menuCommand('mac.data.filter', 'filter'),
        compactCommand('mac.data.clear', 'clear'),
        compactCommand('mac.data.reapply', 'refresh'),
        compactCommand('mac.data.advancedFilter', 'filter'),
      ],
      'sort-filter',
    ),
    group(
      'Data Tools',
      [
        compactCommand('mac.data.textToColumns', 'textToColumns'),
        disabledCompactCommand('mac.data.dataCleanup', 'clear', 'officeRequired'),
        disabledCompactCommand('mac.data.dataAnalysis', 'inspect', 'officeRequired'),
        compactCommand('mac.data.flashFill', 'flashFill'),
        compactCommand('mac.data.removeDuplicates', 'removeDuplicates'),
        compactCommand('mac.data.validation', 'dataValidation'),
        compactCommand('mac.data.consolidate', 'merge'),
      ],
      'compact',
    ),
    group('Forecast', [menuCommand('mac.data.whatIf', 'whatIf')], 'compact'),
    group(
      'Outline',
      [
        compactCommand('mac.data.group', 'outlineGroup'),
        compactCommand('mac.data.ungroup', 'outlineUngroup'),
        compactCommand('mac.data.subtotal', 'outlineGroup'),
        compactCommand('mac.data.showDetail', 'outlineShow'),
        compactCommand('mac.data.hideDetail', 'outlineHide'),
      ],
      'compact',
    ),
    group(
      'Analysis',
      [disabledCompactCommand('mac.data.analysis', 'chart', 'notImplemented')],
      'compact',
    ),
  ]);

const reviewTab = (): MacTabDraft =>
  macTab('review', [
    group('Proofing', [
      menuCommand('mac.review.spelling', 'spelling'),
      disabledCommand('mac.review.thesaurus', 'spelling', 'embeddedUnsupported'),
      menuCommand('mac.review.stats', 'count'),
    ]),
    group('Accessibility', [menuCommand('mac.review.accessibility', 'accessibility')]),
    group('Language', [disabledCommand('mac.review.translate', 'translate', 'notImplemented')]),
    group('Changes', [
      disabledCommand('mac.review.showChanges', 'showChanges', 'embeddedUnsupported'),
    ]),
    group('Comments', [
      menuCommand('mac.review.newComment', 'commentAdd'),
      menuCommand('mac.review.deleteComment', 'clear'),
      menuCommand('mac.review.previousComment', 'goTo'),
      menuCommand('mac.review.nextComment', 'goTo'),
      menuCommand('mac.review.showComments', 'commentAdd'),
    ]),
    group('Notes', [menuCommand('mac.review.notes', 'commentAdd')]),
    group('Protection Changes', [
      menuCommand('mac.review.protectSheet', 'protect'),
      menuCommand('mac.review.protectWorkbook', 'protect'),
    ]),
    group('Protect', [disabledCommand('mac.review.alwaysReadOnly', 'protect', 'notImplemented')]),
    group('Ink', [disabledCommand('mac.review.hideInk', 'pen', 'embeddedUnsupported')]),
  ]);

const viewTab = (): MacTabDraft =>
  macTab('view', [
    group(
      'Sheet Views',
      [
        selectCommand('sheetViewSelect', SHEET_VIEW_OPTIONS),
        compactCommand('mac.view.sheetViewKeep', 'save'),
        compactCommand('mac.view.sheetViewExit', 'paneClose'),
        compactCommand('mac.view.sheetViewNew', 'plus'),
        compactCommand('mac.view.sheetViewOptions', 'options'),
      ],
      'compact',
    ),
    group(
      'Workbook Views',
      [
        compactCommand('mac.view.standard', 'table'),
        compactCommand('mac.view.pageBreakPreview', 'table'),
        compactCommand('mac.view.pageLayout', 'page'),
        disabledCompactCommand('mac.view.customViews', 'customViews', 'notImplemented'),
      ],
      'compact',
    ),
    group(
      'Show',
      [
        disabledCompactCommand('mac.view.navigation', 'show', 'notImplemented'),
        disabledCheckboxCommand('mac.view.ruler', 'notImplemented'),
        checkboxCommand('mac.view.gridlines', 'show'),
        checkboxCommand('mac.view.formulaBar', 'show'),
        checkboxCommand('mac.view.headings', 'show'),
        disabledCompactCommand('mac.view.dataTypeIcons', 'show', 'notImplemented'),
        compactCommand('mac.view.cellFocus', 'show'),
      ],
      'compact',
    ),
    group(
      'Zoom',
      [
        compactCommand('mac.view.zoom', 'zoom'),
        compactCommand('mac.view.zoom100', 'zoom'),
        compactCommand('mac.view.zoomSelection', 'zoom'),
      ],
      'compact',
    ),
    group(
      'Window',
      [
        disabledCompactCommand('mac.view.newWindow', 'window', 'embeddedUnsupported'),
        disabledCompactCommand('mac.view.arrange', 'arrange', 'embeddedUnsupported'),
        compactCommand('mac.view.freeze', 'freeze'),
        compactCommand('mac.view.split', 'split'),
        disabledCompactCommand('mac.view.hide', 'hide', 'embeddedUnsupported'),
        disabledCompactCommand('mac.view.unhide', 'show', 'embeddedUnsupported'),
        disabledCompactCommand('mac.view.sideBySide', 'window', 'embeddedUnsupported'),
        disabledCompactCommand('mac.view.syncScroll', 'window', 'embeddedUnsupported'),
        disabledCompactCommand('mac.view.resetWindow', 'window', 'embeddedUnsupported'),
        disabledCompactCommand('mac.view.switchWindow', 'window', 'embeddedUnsupported'),
      ],
      'compact',
    ),
    group(
      'Macros',
      [
        disabledCompactCommand('mac.view.macros', 'script', 'notImplemented'),
        disabledCompactCommand('mac.view.recordMacro', 'script', 'notImplemented'),
        disabledCompactCommand('mac.view.relativeRecording', 'script', 'notImplemented'),
      ],
      'compact',
    ),
  ]);

const automateTab = (): MacTabDraft =>
  macTab('automate', [
    group('Office Scripts', [
      disabledCommand('mac.automate.newScript', 'script', 'officeRequired'),
      menuCommand('mac.automate.showScripts', 'script'),
    ]),
    group(
      'Office Scripts Gallery',
      [
        compactCommand('mac.automate.allRowsColumns', 'show'),
        compactCommand('mac.automate.freezeSelection', 'freeze'),
        compactCommand('mac.automate.makeSubtable', 'table'),
        compactCommand('mac.automate.removeHyperlinks', 'link'),
        compactCommand('mac.automate.countEmptyRows', 'count'),
        compactCommand('mac.automate.tableToJson', 'code'),
        compactCommand('mac.automate.newPivotTable', 'pivotTable'),
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

/** Mac tooltips are the bare control name: no shortcut or symbol suffix. */
const stripShortcutHint = (title: string): string =>
  title.replace(/\s*\((?:Cmd|Ctrl|⌘)[^)]*\)|\s*\(Σ\)/g, '');

const macHomeCommand = (candidate: RibbonCommand, lang: ToolbarLang): RibbonCommand => {
  // The Mac app's own control name and help text, where they differ from the shared model.
  const measuredName = findMacBundleLabel(`home:${candidate.id}`)?.[lang];
  const measuredHelp = findMacBundleLabel(`home-help:${candidate.id}`)?.[lang];
  const commandWithMacTitle = {
    ...candidate,
    ...(measuredName ? { label: measuredName, title: measuredName } : {}),
    title: stripShortcutHint(measuredName ?? candidate.title),
  };
  const macLabelKey = MAC_HOME_LABEL_KEYS[candidate.id];
  const homeLabel =
    candidate.id === 'merge'
      ? dictionaries[lang].ribbon.mergeAndCenter
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
  // The desktop font-size list starts at 6.
  if (
    candidate.id === 'fontSize' &&
    candidate.options &&
    !candidate.options.some((o) => o.value === '6')
  ) {
    commandWithMacHomeLabel.options = [{ value: '6', label: '6' }, ...candidate.options];
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

/** A shared Home command keeps its own text unless its id has a Mac label. */
const localizeHomeCommand = (candidate: RibbonCommand, lang: ToolbarLang): RibbonCommand => {
  const shaped = macHomeCommand(candidate, lang);
  const label = macRibbonLabelForCommand(shaped.id, lang);
  return label === shaped.id ? shaped : { ...shaped, title: label, label };
};

const macHomeTab = (home: RibbonTabModel, lang: ToolbarLang): RibbonTabModel => ({
  ...home,
  label: labels[home.id]?.[lang] ?? home.label,
  groups: [
    ...home.groups.map((shared, index) => {
      const current = {
        ...shared,
        title: macGroupTitle(MAC_HOME_GROUP_KEYS[index] ?? shared.title, lang),
      };
      if (current.variant === 'cells') {
        // Tile layout rather than the shared stacked-rows layout.
        return {
          ...current,
          variant: 'tiles',
          commands: current.commands.map((candidate) => localizeHomeCommand(candidate, lang)),
        };
      }
      if (current.variant !== 'font') {
        return {
          ...current,
          commands: current.commands.map((candidate) => localizeHomeCommand(candidate, lang)),
        };
      }
      const isPhonetic = (id: string): boolean => id === 'showPhonetic';
      const phonetic = current.commands.filter((candidate) => isPhonetic(candidate.id));
      return {
        ...current,
        commands: [
          ...current.commands
            .filter((candidate) => candidate.id !== 'strike' && !isPhonetic(candidate.id))
            .map((candidate) => localizeHomeCommand(candidate, lang)),
          ...phonetic.map((candidate) => localizeHomeCommand(candidate, lang)),
        ],
      };
    }),
    localizeGroup(group('Add-ins', [menuCommand('mac.home.addins', 'addIn')]), lang),
  ],
});

const localizeTab = (draft: MacTabDraft, lang: ToolbarLang): RibbonTabModel => ({
  id: draft.id,
  label: t(draft.id)[lang],
  groups: draft.groups.map((g) => localizeGroup(g, lang)),
});

/** Builds the Office 365 Mac command surface. The shared Home model is passed
 *  in by ribbon-model.ts so its established formatting controls and active
 *  state projection remain identical across profiles. */
export const buildMacRibbonModel = (lang: ToolbarLang, home: RibbonTabModel): RibbonTabModel[] => [
  macHomeTab(home, lang),
  ...[
    insertTab(),
    drawTab(),
    pageLayoutTab(),
    formulasTab(),
    dataTab(),
    reviewTab(),
    viewTab(),
    automateTab(),
  ].map((tab) => localizeTab(tab, lang)),
];

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
      ...allFunctionNames().map(macFunctionCommandId),
    ]),
  ),
);

export const MAC_RIBBON_COMMAND_ID_SET: ReadonlySet<string> = new Set(MAC_RIBBON_COMMAND_IDS);

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
