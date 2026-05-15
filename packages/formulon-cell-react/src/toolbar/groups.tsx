import {
  bumpDecimals,
  clearFormat,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  type MarginPreset,
  type PageOrientation,
  type PaperSize,
  setAlign,
  setFillColor,
  setFont,
  setFontColor,
  setNumFmt,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '@libraz/formulon-cell';
import type { ReactElement } from 'react';
import { buildAddInRibbonGroups } from './add-in-groups.js';
import type { BuildRibbonGroupsOptions } from './group-types.js';
import { Icon } from './icons.js';
import { FONT_FAMILIES, FONT_SIZES, RIBBON_TAB_LABELS, type RibbonTab } from './model.js';

export const buildRibbonGroups = ({
  active,
  color,
  group,
  iconLabel,
  instance,
  lang,
  mergeMenu,
  onAutoSum,
  onCopy,
  onCut,
  onDeleteCols,
  onDeleteRows,
  onAddIn,
  onFilterToggle,
  onFormatAsTable,
  onFormatPainter,
  onFreezeToggle,
  onDrawEraser,
  onDrawPen,
  onInsertCols,
  onInsertRows,
  onMarginPreset,
  onNumberFormat,
  onPageOrientation,
  onPaperSize,
  onPaste,
  onRedo,
  onRemoveDuplicates,
  onAccessibilityCheck,
  onRunScript,
  onSort,
  onSpellingReview,
  onTranslate,
  onToggleColsHidden,
  onToggleRowsHidden,
  onUndo,
  onZoom,
  optionSelect,
  rowBreak,
  select,
  tool,
  tr,
  wrapFormat,
}: BuildRibbonGroupsOptions): Record<RibbonTab, ReactElement[]> => {
  return {
    file: [
      group(tr.workbook, [
        tool(
          'pageSetup',
          tr.pageSetup,
          iconLabel('page', tr.pageSetup),
          () => instance?.openPageSetup(),
          false,
          ' demo__rb--wide',
        ),
        tool(
          'print',
          tr.print,
          iconLabel('print', tr.print),
          () => instance?.print(),
          false,
          ' demo__rb--wide',
        ),
        tool(
          'links',
          tr.links,
          iconLabel('link', tr.links),
          () => instance?.openExternalLinksDialog(),
          false,
          ' demo__rb--wide',
        ),
      ]),
      group(tr.inspect, [
        tool(
          'formatCells',
          tr.formatCells,
          iconLabel('formatCells', tr.formatCells),
          () => instance?.openFormatDialog(),
          false,
          ' demo__rb--wide',
        ),
        tool(
          'gotoSpecial',
          tr.gotoSpecial,
          iconLabel('goTo', tr.gotoSpecial),
          () => instance?.openGoToSpecial(),
          false,
          ' demo__rb--wide',
        ),
      ]),
    ],
    home: [
      group(
        tr.clipboard,
        [
          tool(
            'paste',
            tr.paste,
            <>
              <Icon name="paste" />
              <span>{tr.paste}</span>
            </>,
            onPaste,
            false,
            ' demo__rb--large',
          ),
          tool('cut', tr.cut, <Icon name="cut" />, onCut),
          tool('copy', tr.copy, <Icon name="copy" />, onCopy),
          tool(
            'formatPainter',
            tr.formatPainter,
            <Icon name="paint" />,
            onFormatPainter,
            active.formatPainterArmed,
          ),
          tool(
            'clearFormat',
            tr.clearFormats,
            <Icon name="clear" />,
            () => wrapFormat(clearFormat),
            false,
            ' demo__rb--wide',
          ),
        ],
        'clipboard',
      ),
      group(
        tr.font,
        [
          select(
            'fontFamily',
            'Font',
            active.fontFamily,
            FONT_FAMILIES,
            (value) => wrapFormat((s, st) => setFont(s, st, { fontFamily: value })),
            ' demo__rb-select--font',
          ),
          select('fontSize', 'Font size', active.fontSize, FONT_SIZES, (value) =>
            wrapFormat((s, st) => setFont(s, st, { fontSize: Number(value) })),
          ),
          tool('fontGrow', tr.increaseFontSize, <Icon name="fontGrow" />, () =>
            wrapFormat((s, st) => setFont(s, st, { fontSize: active.fontSize + 1 })),
          ),
          tool('fontShrink', tr.decreaseFontSize, <Icon name="fontShrink" />, () =>
            wrapFormat((s, st) => setFont(s, st, { fontSize: Math.max(1, active.fontSize - 1) })),
          ),
          rowBreak('font-row-2'),
          tool(
            'bold',
            'Bold (⌘B)',
            <Icon name="bold" />,
            () => wrapFormat(toggleBold),
            active.bold,
            ' demo__rb--bold',
          ),
          tool(
            'italic',
            'Italic (⌘I)',
            <Icon name="italic" />,
            () => wrapFormat(toggleItalic),
            active.italic,
            ' demo__rb--italic',
          ),
          tool(
            'underline',
            'Underline (⌘U)',
            <Icon name="underline" />,
            () => wrapFormat(toggleUnderline),
            active.underline,
            ' demo__rb--underline',
          ),
          tool(
            'strike',
            'Strikethrough',
            <Icon name="strike" />,
            () => wrapFormat(toggleStrike),
            active.strike,
            ' demo__rb--strike',
          ),
          tool('borders', tr.borders, <Icon name="borders" />, () => wrapFormat(cycleBorders)),
          color(
            'fontColor',
            'Font color',
            active.fontColor,
            (value) => wrapFormat((s, st) => setFontColor(s, st, value)),
            <Icon name="fontColor" />,
          ),
          color(
            'fillColor',
            'Fill color',
            active.fillColor,
            (value) => wrapFormat((s, st) => setFillColor(s, st, value)),
            <Icon name="fillColor" />,
          ),
        ],
        'font',
      ),
      group(
        tr.alignment,
        [
          tool(
            'top',
            'Top align',
            <Icon name="top" />,
            () => wrapFormat((s, st) => setVAlign(s, st, 'top')),
            false,
          ),
          tool(
            'middle',
            'Middle align',
            <Icon name="middle" />,
            () => wrapFormat((s, st) => setVAlign(s, st, 'middle')),
            false,
          ),
          rowBreak('alignment-row-2'),
          tool(
            'alignL',
            'Align left',
            <Icon name="alignLeft" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'left')),
            active.alignLeft,
          ),
          tool(
            'alignC',
            'Align center',
            <Icon name="alignCenter" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'center')),
            active.alignCenter,
          ),
          tool(
            'alignR',
            'Align right',
            <Icon name="alignRight" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'right')),
            active.alignRight,
          ),
          tool('wrap', tr.wrapText, <Icon name="wrap" />, () => wrapFormat(toggleWrap)),
          mergeMenu,
        ],
        'alignment',
      ),
      group(
        tr.number,
        [
          optionSelect(
            'numberFormat',
            'Number format',
            active.numberFormat,
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
            onNumberFormat,
            ' demo__rb-select--number-format',
          ),
          rowBreak('number-row-2'),
          tool(
            'currency',
            'Currency',
            <Icon name="currency" />,
            () => wrapFormat(cycleCurrency),
            active.currency,
            ' demo__rb--mono',
          ),
          tool(
            'percent',
            'Percent',
            <Icon name="percent" />,
            () => wrapFormat(cyclePercent),
            active.percent,
            ' demo__rb--mono',
          ),
          tool('comma', tr.commaStyle, <Icon name="comma" />, () =>
            wrapFormat((s, st) => setNumFmt(s, st, { kind: 'fixed', decimals: 2 })),
          ),
          tool('decDown', tr.decreaseDecimals, <Icon name="decDown" />, () =>
            wrapFormat((s, st) => bumpDecimals(s, st, -1)),
          ),
          tool('decUp', tr.increaseDecimals, <Icon name="decUp" />, () =>
            wrapFormat((s, st) => bumpDecimals(s, st, 1)),
          ),
        ],
        'number',
      ),
      group(
        tr.styles,
        [
          tool(
            'conditional',
            'Conditional formatting',
            iconLabel('conditional', tr.conditional),
            () => instance?.openConditionalDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'cellStyles',
            'Cell styles',
            iconLabel('tableStyle', tr.cellStyles),
            () => instance?.openCellStylesGallery(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'rules',
            'Manage conditional formatting rules',
            iconLabel('options', tr.rules),
            () => instance?.openCfRulesDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'styles',
      ),
      group(
        tr.cells,
        [
          tool('insertRows', tr.insertRows, <Icon name="insertRows" />, onInsertRows),
          tool('deleteRows', tr.deleteRows, <Icon name="deleteRows" />, onDeleteRows),
          tool('insertCols', tr.insertCols, <Icon name="insertCols" />, onInsertCols),
          tool('deleteCols', tr.deleteCols, <Icon name="deleteCols" />, onDeleteCols),
          tool(
            'formatCellsHome',
            'Format cells',
            iconLabel('formatCells', tr.formatCells),
            () => instance?.openFormatDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'cells',
      ),
      group(
        tr.editing,
        [
          tool('autosum', 'AutoSum (Σ)', <Icon name="autosum" />, onAutoSum),
          tool('undoHome', 'Undo (⌘Z)', <Icon name="undo" />, onUndo),
          tool('redoHome', 'Redo (⌘⇧Z)', <Icon name="redo" />, onRedo),
          tool('sortAscHome', 'Sort ascending', <Icon name="sortAsc" />, () => onSort('asc')),
          tool('filterHome', 'Filter', <Icon name="filter" />, onFilterToggle, active.filterOn),
          tool(
            'findHome',
            `${tr.find} (⌘F)`,
            iconLabel('find', tr.find),
            () => instance?.openFindReplace(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'gotoSpecialHome',
            tr.gotoSpecial,
            iconLabel('goTo', tr.gotoSpecial),
            () => instance?.openGoToSpecial(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'editing',
      ),
    ],
    insert: [
      group(
        tr.tables,
        [
          tool(
            'pivotTableInsert',
            'PivotTable',
            iconLabel('table', tr.pivotTable),
            () => instance?.openPivotTableDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'formatTableInsert',
            'Format as Table',
            iconLabel('tableStyle', tr.formatTable),
            onFormatAsTable,
            false,
            ' demo__rb--wide',
          ),
          tool(
            'namedRangesInsert',
            'Name manager',
            iconLabel('names', tr.names),
            () => instance?.openNamedRangeDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'removeDupesInsert',
            'Remove duplicates',
            iconLabel('removeDuplicates', tr.removeDuplicates),
            onRemoveDuplicates,
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.charts,
        [
          tool(
            'chartInsert',
            'Recommended chart',
            iconLabel('chart', tr.chart),
            () => instance?.openQuickAnalysis(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.links,
        [
          tool(
            'hyperlinkInsert',
            'Insert hyperlink (⌘K)',
            iconLabel('link', tr.hyperlink),
            () => instance?.openHyperlinkDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'linksInsert',
            'Edit links',
            iconLabel('link', tr.links),
            () => instance?.openExternalLinksDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.comments,
        [
          tool(
            'commentInsert',
            active.hasComment ? 'Edit Note' : 'New Note',
            iconLabel(
              active.hasComment ? 'commentMultiple' : 'commentAdd',
              active.hasComment ? tr.editComment : tr.newComment,
            ),
            () => instance?.openCommentDialog(),
            active.hasComment,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.symbols,
        [
          tool(
            'fxInsert',
            'Insert function (Σ)',
            iconLabel('function', 'fx'),
            () => instance?.openFunctionArguments(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    draw: [
      group(
        RIBBON_TAB_LABELS.draw[lang],
        [
          tool(
            'drawPen',
            RIBBON_TAB_LABELS.draw[lang],
            iconLabel('pen', tr.pen),
            () => onDrawPen?.(),
            false,
            ' demo__rb--wide',
            !onDrawPen,
          ),
          tool(
            'drawErase',
            'Eraser',
            iconLabel('eraser', tr.eraser),
            () => onDrawEraser?.(),
            false,
            ' demo__rb--wide',
            !onDrawEraser,
          ),
        ],
        'tiles',
      ),
    ],
    pageLayout: [
      group(
        tr.pageSetup,
        [
          optionSelect<MarginPreset | 'custom'>(
            'marginsPreset',
            tr.margins,
            active.marginPreset ?? 'custom',
            [
              { value: 'normal', label: tr.marginsNormal },
              { value: 'wide', label: tr.marginsWide },
              { value: 'narrow', label: tr.marginsNarrow },
              // "Custom" is read-only — selecting it would have to round-trip
              // through Page Setup. We include it so the closed display can
              // honestly say "Custom" when the user has bespoke margins.
              { value: 'custom', label: tr.marginsCustom },
            ],
            (next) => {
              if (next === 'custom') {
                instance?.openPageSetup();
                return;
              }
              onMarginPreset(next);
            },
            ' demo__rb-select--border',
          ),
          optionSelect(
            'orientationPreset',
            tr.orientation,
            active.pageOrientation,
            [
              { value: 'portrait' as PageOrientation, label: tr.portrait },
              { value: 'landscape' as PageOrientation, label: tr.landscape },
            ],
            onPageOrientation,
            ' demo__rb-select--border',
          ),
          optionSelect(
            'paperSizePreset',
            tr.paperSize,
            active.paperSize,
            [
              { value: 'A4' as PaperSize, label: tr.paperA4 },
              { value: 'A3' as PaperSize, label: tr.paperA3 },
              { value: 'A5' as PaperSize, label: tr.paperA5 },
              { value: 'letter' as PaperSize, label: tr.paperLetter },
              { value: 'legal' as PaperSize, label: tr.paperLegal },
              { value: 'tabloid' as PaperSize, label: tr.paperTabloid },
            ],
            onPaperSize,
            ' demo__rb-select--border',
          ),
          tool(
            'pageSetupAdvanced',
            'Advanced page setup',
            iconLabel('options', tr.pageSetup),
            () => instance?.openPageSetup(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.print,
        [
          tool(
            'printPageLayout',
            tr.print,
            iconLabel('print', tr.print),
            () => instance?.print(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    formulas: [
      group(
        tr.functionLibrary,
        [
          tool(
            'fx',
            'Insert function',
            <Icon name="function" />,
            () => instance?.openFunctionArguments(),
            false,
            ' demo__rb--mono',
          ),
          tool(
            'autosumFormula',
            'AutoSum (Σ)',
            <>
              <Icon name="autosum" />
              <span>{tr.autoSum}</span>
            </>,
            onAutoSum,
          ),
          tool(
            'sum',
            'SUM arguments',
            iconLabel('function', 'SUM'),
            () => instance?.openFunctionArguments('SUM'),
            false,
            ' demo__rb--mono',
          ),
          tool(
            'avg',
            'AVERAGE arguments',
            iconLabel('function', 'AVG'),
            () => instance?.openFunctionArguments('AVERAGE'),
            false,
            ' demo__rb--mono',
          ),
        ],
        'tiles',
      ),
      group(
        tr.definedNames,
        [
          tool(
            'namedRanges',
            'Name manager',
            iconLabel('names', tr.names),
            () => instance?.openNamedRangeDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.formulaAuditing,
        [
          tool(
            'precedents',
            'Trace precedents',
            iconLabel('trace', tr.tracePrecedents),
            () => instance?.tracePrecedents(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'dependents',
            'Trace dependents',
            iconLabel('dependents', tr.traceDependents),
            () => instance?.traceDependents(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'clearArrows',
            'Remove arrows',
            iconLabel('clearArrows', tr.removeArrows),
            () => instance?.clearTraces(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.calculation,
        [
          tool(
            'recalcNow',
            'Calculate Now (F9)',
            iconLabel('autosum', tr.recalc),
            () => instance?.recalc(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'calcOptions',
            'Calculation options',
            iconLabel('options', tr.options),
            () => instance?.openIterativeDialog(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'watch',
            'Watch Window',
            iconLabel('watch', tr.watch),
            () => instance?.toggleWatchWindow(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    data: [
      group(
        tr.sortFilter,
        [
          tool(
            'filter',
            'Filter',
            <>
              <Icon name="filter" />
              <span>{tr.filter}</span>
            </>,
            onFilterToggle,
            active.filterOn,
          ),
          tool(
            'sortAsc',
            'Sort ascending',
            <>
              <Icon name="sortAsc" />
              <span>A-Z</span>
            </>,
            () => onSort('asc'),
          ),
          tool(
            'sortDesc',
            'Sort descending',
            <>
              <Icon name="sortDesc" />
              <span>Z-A</span>
            </>,
            () => onSort('desc'),
          ),
        ],
        'tiles',
      ),
      group(
        tr.dataTools,
        [
          tool(
            'removeDupes',
            'Remove duplicates',
            iconLabel('removeDuplicates', tr.removeDuplicates),
            onRemoveDuplicates,
            false,
            ' demo__rb--wide',
          ),
          tool(
            'linksData',
            'Edit links',
            iconLabel('link', tr.links),
            () => instance?.openExternalLinksDialog(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.outline,
        [
          tool(
            'hideRows',
            active.rowsHidden ? 'Show selected rows' : 'Hide selected rows',
            iconLabel('table', active.rowsHidden ? tr.showRows : tr.hideRows),
            onToggleRowsHidden,
            active.rowsHidden,
            ' demo__rb--wide',
          ),
          tool(
            'hideCols',
            active.colsHidden ? 'Show selected columns' : 'Hide selected columns',
            iconLabel('table', active.colsHidden ? tr.showCols : tr.hideCols),
            onToggleColsHidden,
            active.colsHidden,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    review: [
      group(
        tr.proofing,
        [
          tool(
            'spellingReview',
            tr.spelling,
            iconLabel('spelling', tr.spelling),
            () => onSpellingReview?.(),
            false,
            ' demo__rb--wide',
            !onSpellingReview,
          ),
        ],
        'tiles',
      ),
      group(
        tr.language,
        [
          tool(
            'translateReview',
            tr.translate,
            iconLabel('translate', tr.translate),
            () => onTranslate?.(),
            false,
            ' demo__rb--wide',
            !onTranslate,
          ),
        ],
        'tiles',
      ),
      group(
        tr.comments,
        [
          tool(
            'newCommentReview',
            active.hasComment ? 'Edit Note' : 'New Note',
            iconLabel(
              active.hasComment ? 'commentMultiple' : 'commentAdd',
              active.hasComment ? tr.editComment : tr.newComment,
            ),
            () => instance?.openCommentDialog(),
            active.hasComment,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.find,
        [
          tool(
            'findReview',
            `${tr.find} (⌘F)`,
            iconLabel('find', tr.find),
            () => instance?.openFindReplace(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.protection,
        [
          tool(
            'protectReview',
            active.protected ? 'Unprotect sheet' : 'Protect sheet',
            iconLabel('protect', active.protected ? tr.unprotect : tr.protect),
            () => instance?.toggleSheetProtection(),
            active.protected,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.accessibility,
        [
          tool(
            'accessibility',
            tr.accessibility,
            iconLabel('accessibility', tr.accessibility),
            () => onAccessibilityCheck?.(),
            false,
            ' demo__rb--wide',
            !onAccessibilityCheck,
          ),
        ],
        'tiles',
      ),
    ],
    view: [
      group(
        tr.workbookViews,
        [
          tool(
            'watchView',
            'Watch Window',
            iconLabel('watch', tr.watch),
            () => instance?.toggleWatchWindow(),
            false,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.window,
        [
          tool(
            'freeze',
            'Freeze panes',
            <>
              <Icon name="freeze" />
              <span>{tr.freeze}</span>
            </>,
            onFreezeToggle,
            active.frozen,
          ),
        ],
        'tiles',
      ),
      group(
        tr.zoom,
        [
          tool(
            'zoom75',
            'Zoom to 75%',
            iconLabel('zoom', '75%'),
            () => onZoom(0.75),
            active.zoom === 0.75,
            ' demo__rb--mono',
          ),
          tool(
            'zoom100',
            'Zoom to 100%',
            iconLabel('zoom', '100%'),
            () => onZoom(1),
            active.zoom === 1,
            ' demo__rb--mono',
          ),
          tool(
            'zoom125',
            'Zoom to 125%',
            iconLabel('zoom', '125%'),
            () => onZoom(1.25),
            active.zoom === 1.25,
            ' demo__rb--mono',
          ),
        ],
        'tiles',
      ),
      group(
        tr.protection,
        [
          tool(
            'protect',
            active.protected ? 'Unprotect sheet' : 'Protect sheet',
            iconLabel('protect', active.protected ? tr.unprotect : tr.protect),
            () => instance?.toggleSheetProtection(),
            active.protected,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    ...buildAddInRibbonGroups({
      group,
      iconLabel,
      instance,
      lang,
      onAddIn,
      onRunScript,
      tool,
      tr,
    }),
  };
};
