import {
  activateSheetView,
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  buildTranslationReviewItems,
  bumpDecimals,
  bumpIndent,
  collapseColGroup,
  collapseRowGroup,
  cycleBorders,
  cycleCurrency,
  cyclePercent,
  deleteSheetView,
  groupCols,
  groupRows,
  listComments,
  type MarginPreset,
  mutators,
  type PageOrientation,
  type PaperSize,
  reviewCellsFromState,
  saveSheetView,
  setAlign,
  setFillColor,
  setFont,
  setFontColor,
  setGridlinesVisible,
  setHeadingsVisible,
  setNumFmt,
  setPrintGridlines,
  setPrintHeadings,
  setR1C1ReferenceStyle,
  setShowFormulas,
  setVAlign,
  setWorkbookView,
  showCols,
  showRows,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
  ungroupCols,
  ungroupRows,
} from '@libraz/formulon-cell';
import type { ReactElement } from 'react';
import { buildAddInRibbonGroups } from './add-in-groups.js';
import type { BuildRibbonGroupsOptions } from './group-types.js';
import { Icon } from './icons.js';
import { FONT_SIZES, type RibbonTab } from './model.js';

export const buildRibbonGroups = ({
  active,
  addInMenu,
  autosumFormulaMenu,
  autosumMenu,
  borderPresets,
  borderColor,
  borderStyle,
  borderStyles,
  cellDeleteMenu,
  cellFormatMenu,
  cellInsertMenu,
  cellStylesMenu,
  calcOptionsMenu,
  chartMenu,
  clearMenu,
  clearArrowsMenu,
  color,
  conditionalMenu,
  dataFilterMenu,
  dataSortMenu,
  dataValidationMenu,
  definedNamesMenu,
  deleteCommentMenu,
  errorCheckingMenu,
  freezeMenu,
  windowMenu,
  formatTableHomeMenu,
  formatTableInsertMenu,
  fillMenu,
  findMenu,
  functionDateTimeMenu,
  functionFinancialMenu,
  functionLogicalMenu,
  functionLookupMenu,
  functionMathTrigMenu,
  functionTextMenu,
  hyperlinkMenu,
  formulaBarVisible,
  group,
  iconLabel,
  instance,
  lang,
  strings,
  mergeMenu,
  onCopy,
  onCut,
  onFormatPainter,
  onDrawEraser,
  onDrawPen,
  onMarginPreset,
  onNumberFormat,
  onPageOrientation,
  onPaperSize,
  onProtectWorkbook,
  onInspectWorkbook,
  onRemoveDuplicates,
  onScaleFit,
  onScalePercent,
  onAccessibilityCheck,
  onBorderPreset,
  onRunScript,
  onRecordActions,
  onAllScripts,
  onBuiltInReview,
  onSort,
  onSpellingReview,
  onTranslate,
  onToggleFormulaBar,
  onZoom,
  onZoomDialog,
  onZoomSelection,
  pdfMenu,
  pivotTableMenu,
  protectionMenu,
  optionSelect,
  rowBreak,
  select,
  setBorderStyle,
  setBorderColor,
  printAreaMenu,
  pageBreaksMenu,
  pasteMenu,
  outlineGroupMenu,
  outlineUngroupMenu,
  pictureInsertMenu,
  sheetBackgroundMenu,
  shapesInsertMenu,
  screenshotInsertMenu,
  printTitlesMenu,
  sortMenu,
  symbolMenu,
  themeMenu,
  textOrientationMenu,
  textToColumnsMenu,
  watchMenu,
  watchViewMenu,
  tool,
  tr,
  workbookStructureProtected,
  wrapFormat,
}: BuildRibbonGroupsOptions): Record<RibbonTab, ReactElement[]> => {
  const pageScaleText = strings.pageScale;
  const scalePageOption = (value: '1' | '2' | '3'): string =>
    `${value} ${value === '1' ? pageScaleText.page : pageScaleText.pages}`;
  const viewText = strings.viewToggle;
  const pageSetupText = strings.pageSetup;
  const cellMenuText = strings.ribbonMenu;
  const functionArgsTitle = (name: string): string =>
    strings.ribbon.functionArgumentsTitle.replace('{name}', name);
  const showBuiltInReview = (
    title: string,
    items: ReturnType<typeof analyzeSpellingCells>,
  ): void => {
    onBuiltInReview?.(title, items);
  };
  const runSpellingReview = (): void => {
    if (onSpellingReview) {
      onSpellingReview();
      return;
    }
    if (!instance) return;
    showBuiltInReview(
      tr.spelling,
      analyzeSpellingCells(reviewCellsFromState(instance.store.getState()), lang),
    );
  };
  const runAccessibilityReview = (): void => {
    if (onAccessibilityCheck) {
      onAccessibilityCheck();
      return;
    }
    if (!instance) return;
    showBuiltInReview(
      tr.accessibility,
      analyzeAccessibilityCells(reviewCellsFromState(instance.store.getState()), lang),
    );
  };
  const runTranslateReview = (): void => {
    if (onTranslate) {
      onTranslate();
      return;
    }
    if (!instance) return;
    const state = instance.store.getState();
    showBuiltInReview(
      tr.translate,
      buildTranslationReviewItems(
        reviewCellsFromState(state, state.data.sheetIndex, state.selection.range),
        lang,
      ),
    );
  };
  const runDrawPen = (): void => {
    if (onDrawPen) {
      onDrawPen();
      return;
    }
    instance?.borderDraw?.activate('draw', borderStyle, borderColor);
  };
  const runDrawGrid = (): void => {
    instance?.borderDraw?.activate('grid', borderStyle, borderColor);
  };
  const runDrawEraser = (): void => {
    if (onDrawEraser) {
      onDrawEraser();
      return;
    }
    instance?.borderDraw?.activate('erase');
  };
  const selectComment = (direction: 1 | -1): void => {
    if (!instance) return;
    const state = instance.store.getState();
    const comments = listComments(state);
    if (comments.length === 0) return;
    const activeAddr = state.selection.active;
    const current = comments.findIndex(
      (entry) => entry.addr.row === activeAddr.row && entry.addr.col === activeAddr.col,
    );
    const nextIndex =
      current >= 0
        ? (current + direction + comments.length) % comments.length
        : direction > 0
          ? 0
          : comments.length - 1;
    const next = comments[nextIndex]?.addr;
    if (next) mutators.setActive(instance.store, next);
  };
  const selectionOutlineAxis = (): 'row' | 'col' => {
    if (!instance) return 'row';
    const range = instance.store.getState().selection.range;
    return range.r1 - range.r0 >= range.c1 - range.c0 ? 'row' : 'col';
  };
  const sheetViewOptions = (): readonly { value: string; label: string }[] => {
    const state = instance?.store.getState();
    const current = { value: 'current', label: strings.viewToolbar.currentView };
    if (!state) return [current];
    return [
      current,
      ...state.sheetViews.views
        .filter((view) => view.sheet === state.data.sheetIndex)
        .map((view) => ({ value: view.id, label: view.name })),
    ];
  };
  const saveSheetViewFromRibbon = (): void => {
    if (!instance) return;
    const count = instance.store.getState().sheetViews.views.length + 1;
    const id = `view-${Date.now().toString(36)}-${count}`;
    saveSheetView(instance.store, id, `${strings.viewToolbar.views} ${count}`);
    instance.store.setState((state) => ({
      ...state,
      sheetViews: { ...state.sheetViews, activeViewId: id },
    }));
  };
  const deleteActiveSheetViewFromRibbon = (): void => {
    if (!instance) return;
    const id = instance.store.getState().sheetViews.activeViewId;
    if (id) deleteSheetView(instance.store, id);
  };
  const applyOutlineAction = (
    action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail',
  ): void => {
    if (!instance) return;
    const range = instance.store.getState().selection.range;
    if (selectionOutlineAxis() === 'row') {
      if (action === 'group')
        groupRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      else if (action === 'ungroup')
        ungroupRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      else if (action === 'show-detail')
        showRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      else
        collapseRowGroup(instance.store, instance.history, range.r0, range.r1, instance.workbook);
    } else {
      if (action === 'group')
        groupCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
      else if (action === 'ungroup')
        ungroupCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
      else if (action === 'show-detail')
        showCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
      else
        collapseColGroup(instance.store, instance.history, range.c0, range.c1, instance.workbook);
    }
  };
  const toggleViewFlag = (flag: 'gridlines' | 'headings' | 'formulas' | 'r1c1'): void => {
    if (!instance) return;
    const ui = instance.store.getState().ui;
    if (flag === 'gridlines') setGridlinesVisible(instance.store, ui.showGridLines === false);
    else if (flag === 'headings') setHeadingsVisible(instance.store, ui.showHeaders === false);
    else if (flag === 'formulas') setShowFormulas(instance.store, !ui.showFormulas);
    else setR1C1ReferenceStyle(instance.store, !ui.r1c1);
  };
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
          () => instance?.print('print'),
          false,
          ' demo__rb--wide',
        ),
        tool(
          'protect',
          tr.protect,
          iconLabel('protect', tr.protect),
          () => onProtectWorkbook?.(),
          false,
          ' demo__rb--wide',
        ),
      ]),
      group(tr.inspect, [
        tool(
          'inspect',
          tr.inspect,
          iconLabel('goTo', tr.inspect),
          () => onInspectWorkbook?.(),
          false,
          ' demo__rb--wide',
        ),
      ]),
    ],
    home: [
      group(
        tr.clipboard,
        [
          pasteMenu,
          tool('cut', tr.cut, <Icon name="cut" />, onCut),
          tool('copy', tr.copy, <Icon name="copy" />, onCopy),
          tool(
            'formatPainter',
            tr.formatPainter,
            <Icon name="paint" />,
            onFormatPainter,
            active.formatPainterArmed,
          ),
        ],
        'clipboard',
      ),
      group(
        tr.font,
        [
          select(
            'fontFamily',
            tr.font,
            active.fontFamily,
            strings.ribbon.fontFamilies,
            (value) => wrapFormat((s, st) => setFont(s, st, { fontFamily: value })),
            ' demo__rb-select--font',
          ),
          select('fontSize', tr.fontSize, active.fontSize, FONT_SIZES, (value) =>
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
            `${tr.bold} (⌘B)`,
            <Icon name="bold" />,
            () => wrapFormat(toggleBold),
            active.bold,
            ' demo__rb--bold',
          ),
          tool(
            'italic',
            `${tr.italic} (⌘I)`,
            <Icon name="italic" />,
            () => wrapFormat(toggleItalic),
            active.italic,
            ' demo__rb--italic',
          ),
          tool(
            'underline',
            `${tr.underline} (⌘U)`,
            <Icon name="underline" />,
            () => wrapFormat(toggleUnderline),
            active.underline,
            ' demo__rb--underline',
          ),
          tool(
            'strike',
            tr.strikethrough,
            <Icon name="strike" />,
            () => wrapFormat(toggleStrike),
            active.strike,
            ' demo__rb--strike',
          ),
          tool('borders', tr.borders, <Icon name="borders" />, () => wrapFormat(cycleBorders)),
          optionSelect(
            'borderPreset',
            tr.borderPattern,
            'outline',
            borderPresets,
            onBorderPreset,
            ' demo__rb-select--border',
          ),
          optionSelect(
            'borderStyle',
            tr.borderLineStyle,
            borderStyle,
            borderStyles,
            setBorderStyle,
            ' demo__rb-select--border-style',
          ),
          color(
            'borderColor',
            tr.lineColor,
            borderColor,
            setBorderColor,
            <Icon name="fontColor" />,
          ),
          tool('moreBorders', tr.moreBorders, <Icon name="formatCells" />, () =>
            instance?.openFormatDialog('border'),
          ),
          tool(
            'drawBorder',
            tr.drawBorder,
            <Icon name="pen" />,
            runDrawPen,
            false,
            '',
            !onDrawPen && !instance?.borderDraw,
            !!onDrawPen,
          ),
          tool(
            'drawBorderGrid',
            tr.drawBorderGrid,
            <Icon name="borders" />,
            runDrawGrid,
            false,
            '',
            !instance?.borderDraw,
          ),
          tool(
            'eraseBorder',
            tr.eraseBorder,
            <Icon name="eraser" />,
            runDrawEraser,
            false,
            '',
            !onDrawEraser && !instance?.borderDraw,
            !!onDrawEraser,
          ),
          color(
            'fontColor',
            tr.fontColor,
            active.fontColor,
            (value) => wrapFormat((s, st) => setFontColor(s, st, value)),
            <Icon name="fontColor" />,
          ),
          color(
            'fillColor',
            tr.fillColor,
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
            tr.topAlign,
            <Icon name="top" />,
            () => wrapFormat((s, st) => setVAlign(s, st, 'top')),
            active.vAlignTop,
          ),
          tool(
            'middle',
            tr.middleAlign,
            <Icon name="middle" />,
            () => wrapFormat((s, st) => setVAlign(s, st, 'middle')),
            active.vAlignMiddle,
          ),
          tool(
            'bottomAlign',
            tr.bottomAlign,
            iconLabel('bottomAlign', tr.bottomAlign),
            () => wrapFormat((s, st) => setVAlign(s, st, 'bottom')),
            active.vAlignBottom,
          ),
          textOrientationMenu,
          tool(
            'wrap',
            tr.wrapText,
            <Icon name="wrap" />,
            () => wrapFormat(toggleWrap),
            active.wrapText,
          ),
          rowBreak('alignment-row-2'),
          tool(
            'alignL',
            tr.alignLeft,
            <Icon name="alignLeft" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'left')),
            active.alignLeft,
          ),
          tool(
            'alignC',
            tr.alignCenter,
            <Icon name="alignCenter" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'center')),
            active.alignCenter,
          ),
          tool(
            'alignR',
            tr.alignRight,
            <Icon name="alignRight" />,
            () => wrapFormat((s, st) => setAlign(s, st, 'right')),
            active.alignRight,
          ),
          tool(
            'indentDecrease',
            tr.decreaseIndent,
            iconLabel('indentDecrease', tr.decreaseIndent),
            () => wrapFormat((s, st) => bumpIndent(s, st, -1)),
          ),
          tool(
            'indentIncrease',
            tr.increaseIndent,
            iconLabel('indentIncrease', tr.increaseIndent),
            () => wrapFormat((s, st) => bumpIndent(s, st, 1)),
          ),
          mergeMenu,
        ],
        'alignment',
      ),
      group(
        tr.number,
        [
          optionSelect(
            'numberFormat',
            tr.number,
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
            tr.currency,
            <Icon name="currency" />,
            () => wrapFormat((s, st) => cycleCurrency(s, st, lang)),
            active.currency,
            ' demo__rb--mono',
          ),
          tool(
            'percent',
            tr.percent,
            <Icon name="percent" />,
            () => wrapFormat(cyclePercent),
            active.percent,
            ' demo__rb--mono',
          ),
          tool(
            'comma',
            tr.commaStyle,
            <Icon name="comma" />,
            () =>
              wrapFormat((s, st) =>
                setNumFmt(s, st, { kind: 'fixed', decimals: 2, thousands: true }),
              ),
            active.commaStyle,
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
      group(tr.styles, [conditionalMenu, formatTableHomeMenu, cellStylesMenu], 'styles'),
      group(tr.cells, [cellInsertMenu, cellDeleteMenu, cellFormatMenu], 'cells'),
      group(tr.editing, [autosumMenu, fillMenu, clearMenu, sortMenu, findMenu], 'editing'),
    ],
    insert: [
      group(tr.tables, [pivotTableMenu, formatTableInsertMenu], 'tiles'),
      group(tr.illustrations, [pictureInsertMenu, shapesInsertMenu, screenshotInsertMenu], 'tiles'),
      group(tr.charts, [chartMenu], 'tiles'),
      group(tr.links, [hyperlinkMenu], 'tiles'),
      group(
        tr.comments,
        [
          tool(
            'commentInsert',
            active.hasComment ? tr.editComment : tr.newComment,
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
      group(tr.symbols, [symbolMenu], 'tiles'),
    ],
    draw: [
      group(
        strings.ribbon.tabs.draw,
        [
          tool(
            'drawPen',
            strings.ribbon.tabs.draw,
            iconLabel('pen', tr.pen),
            runDrawPen,
            false,
            ' demo__rb--wide',
            !onDrawPen && !instance?.borderDraw,
            !!onDrawPen,
          ),
          tool(
            'drawGrid',
            tr.drawBorderGrid,
            iconLabel('borders', tr.drawBorderGrid),
            runDrawGrid,
            false,
            ' demo__rb--wide',
            !instance?.borderDraw,
          ),
          tool(
            'drawErase',
            tr.eraser,
            iconLabel('eraser', tr.eraser),
            runDrawEraser,
            false,
            ' demo__rb--wide',
            !onDrawEraser && !instance?.borderDraw,
            !!onDrawEraser,
          ),
        ],
        'tiles',
      ),
    ],
    pageLayout: [
      group(cellMenuText.theme, [themeMenu], 'tiles'),
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
            tr.pageSetup,
            iconLabel('options', tr.pageSetup),
            () => instance?.openPageSetup(),
            false,
            ' demo__rb--wide',
          ),
          printAreaMenu,
          pageBreaksMenu,
          sheetBackgroundMenu,
          printTitlesMenu,
        ],
        'tiles',
      ),
      group(
        tr.scale,
        [
          optionSelect(
            'scaleWidth',
            pageScaleText.width,
            active.fitWidth == null ? '0' : String(active.fitWidth),
            [
              { value: '0', label: pageScaleText.automatic },
              { value: '1', label: scalePageOption('1') },
              { value: '2', label: scalePageOption('2') },
              { value: '3', label: scalePageOption('3') },
            ],
            (value) => onScaleFit('width', value),
            ' demo__rb-select--border',
          ),
          optionSelect(
            'scaleHeight',
            pageScaleText.height,
            active.fitHeight == null ? '0' : String(active.fitHeight),
            [
              { value: '0', label: pageScaleText.automatic },
              { value: '1', label: scalePageOption('1') },
              { value: '2', label: scalePageOption('2') },
              { value: '3', label: scalePageOption('3') },
            ],
            (value) => onScaleFit('height', value),
            ' demo__rb-select--border',
          ),
          optionSelect(
            'scalePercent',
            pageScaleText.scale,
            String(Math.round(active.pageScale * 100)),
            [
              { value: '25', label: '25%' },
              { value: '50', label: '50%' },
              { value: '75', label: '75%' },
              { value: '100', label: '100%' },
              { value: '125', label: '125%' },
              { value: '150', label: '150%' },
              { value: '200', label: '200%' },
              { value: '400', label: '400%' },
            ],
            onScalePercent,
            ' demo__rb-select--border',
          ),
        ],
        'tiles',
      ),
      group(
        tr.sheetOptions,
        [
          tool(
            'pageLayoutGridlinesView',
            `${viewText.gridlines} ${tr.show}`,
            iconLabel('table', viewText.gridlines),
            () => toggleViewFlag('gridlines'),
            active.gridlinesVisible,
            ' demo__rb--wide',
          ),
          tool(
            'pageLayoutGridlinesPrint',
            pageSetupText.showGridlines,
            iconLabel('print', pageSetupText.showGridlines),
            () => {
              if (!instance) return;
              const sheet = instance.store.getState().data.sheetIndex;
              setPrintGridlines(instance.store, sheet, !active.printGridlines, instance.history);
            },
            active.printGridlines,
            ' demo__rb--wide',
          ),
          tool(
            'pageLayoutHeadingsView',
            `${viewText.headings} ${tr.show}`,
            iconLabel('table', viewText.headings),
            () => toggleViewFlag('headings'),
            active.headingsVisible,
            ' demo__rb--wide',
          ),
          tool(
            'pageLayoutHeadingsPrint',
            pageSetupText.showHeadings,
            iconLabel('print', pageSetupText.showHeadings),
            () => {
              if (!instance) return;
              const sheet = instance.store.getState().data.sheetIndex;
              setPrintHeadings(instance.store, sheet, !active.printHeadings, instance.history);
            },
            active.printHeadings,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(
        tr.arrange,
        [
          tool(
            'arrangeObjectsPageLayout',
            tr.arrange,
            iconLabel('options', tr.arrange),
            () => instance?.openWorkbookObjects(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'selectionPanePageLayout',
            tr.selectionPane,
            iconLabel('options', tr.selectionPane),
            () => instance?.openWorkbookObjects(),
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
            () => instance?.print('print'),
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
            tr.insertFunction,
            <Icon name="function" />,
            () => instance?.openFunctionArguments(),
            false,
            ' demo__rb--mono',
          ),
          autosumFormulaMenu,
          tool(
            'sum',
            functionArgsTitle('SUM'),
            iconLabel('function', 'SUM'),
            () => instance?.openFunctionArguments('SUM'),
            false,
            ' demo__rb--mono',
          ),
          tool(
            'avg',
            functionArgsTitle('AVERAGE'),
            iconLabel('function', 'AVG'),
            () => instance?.openFunctionArguments('AVERAGE'),
            false,
            ' demo__rb--mono',
          ),
          functionLogicalMenu,
          functionLookupMenu,
          functionTextMenu,
          functionDateTimeMenu,
          functionFinancialMenu,
          functionMathTrigMenu,
        ],
        'tiles',
      ),
      group(tr.definedNames, [definedNamesMenu], 'tiles'),
      group(
        tr.formulaAuditing,
        [
          tool(
            'precedents',
            tr.tracePrecedents,
            iconLabel('trace', tr.tracePrecedents),
            () => instance?.tracePrecedents(),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'dependents',
            tr.traceDependents,
            iconLabel('dependents', tr.traceDependents),
            () => instance?.traceDependents(),
            false,
            ' demo__rb--wide',
          ),
          clearArrowsMenu,
          errorCheckingMenu,
          tool(
            'showFormulasFormula',
            viewText.formulas,
            iconLabel('function', viewText.formulas),
            () => toggleViewFlag('formulas'),
            active.formulasVisible,
            ' demo__rb--wide',
          ),
          tool(
            'evaluateFormula',
            tr.evaluateFormula,
            iconLabel('function', tr.evaluateFormula),
            () => instance?.openEvaluateFormulaDialog(),
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
            `${tr.recalc} (F9)`,
            iconLabel('autosum', tr.recalc),
            () => instance?.recalc(),
            false,
            ' demo__rb--wide',
          ),
          calcOptionsMenu,
          watchMenu,
        ],
        'tiles',
      ),
    ],
    data: [
      group(
        tr.sortFilter,
        [
          dataFilterMenu,
          tool(
            'sortAsc',
            tr.sortAscending,
            <>
              <Icon name="sortAsc" />
              <span>A-Z</span>
            </>,
            () => onSort('asc'),
          ),
          tool(
            'sortDesc',
            tr.sortDescending,
            <>
              <Icon name="sortDesc" />
              <span>Z-A</span>
            </>,
            () => onSort('desc'),
          ),
          dataSortMenu,
        ],
        'tiles',
      ),
      group(
        tr.dataTools,
        [
          textToColumnsMenu,
          tool(
            'removeDupes',
            tr.removeDuplicates,
            iconLabel('removeDuplicates', tr.removeDuplicates),
            onRemoveDuplicates,
            false,
            ' demo__rb--wide',
          ),
          dataValidationMenu,
          tool(
            'linksData',
            tr.links,
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
          outlineGroupMenu,
          outlineUngroupMenu,
          tool(
            'outlineShowDetail',
            tr.showDetail,
            iconLabel('table', tr.showDetail),
            () => applyOutlineAction('show-detail'),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'outlineHideDetail',
            tr.hideDetail,
            iconLabel('table', tr.hideDetail),
            () => applyOutlineAction('hide-detail'),
            false,
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
            runSpellingReview,
            false,
            ' demo__rb--wide',
            !instance && !onSpellingReview,
            !!onSpellingReview,
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
            runAccessibilityReview,
            false,
            ' demo__rb--wide',
            !instance && !onAccessibilityCheck,
            !!onAccessibilityCheck,
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
            runTranslateReview,
            false,
            ' demo__rb--wide',
            !instance && !onTranslate,
            !!onTranslate,
          ),
        ],
        'tiles',
      ),
      group(
        tr.comments,
        [
          tool(
            'newCommentReview',
            active.hasComment ? tr.editComment : tr.newComment,
            iconLabel(
              active.hasComment ? 'commentMultiple' : 'commentAdd',
              active.hasComment ? tr.editComment : tr.newComment,
            ),
            () => instance?.openCommentDialog(),
            active.hasComment,
            ' demo__rb--wide',
          ),
          deleteCommentMenu,
          tool(
            'previousCommentReview',
            tr.previousComment,
            iconLabel('goTo', tr.previousComment),
            () => selectComment(-1),
            false,
            ' demo__rb--wide',
          ),
          tool(
            'nextCommentReview',
            tr.nextComment,
            iconLabel('goTo', tr.nextComment),
            () => selectComment(1),
            false,
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
            () => instance?.openFindReplace('find'),
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
            active.protected ? tr.unprotect : tr.protect,
            iconLabel('protect', active.protected ? tr.unprotect : tr.protect),
            () => instance?.toggleSheetProtection(),
            active.protected,
            ' demo__rb--wide',
          ),
          tool(
            'protectWorkbookReview',
            workbookStructureProtected
              ? cellMenuText.unprotectWorkbookCommand
              : cellMenuText.protectWorkbookCommand,
            iconLabel(
              'protect',
              workbookStructureProtected
                ? cellMenuText.unprotectWorkbookCommand
                : cellMenuText.protectWorkbookCommand,
            ),
            () => onProtectWorkbook?.(),
            workbookStructureProtected,
            ' demo__rb--wide',
          ),
          protectionMenu,
        ],
        'tiles',
      ),
    ],
    view: [
      group(
        tr.workbookViews,
        [
          tool(
            'viewNormal',
            tr.normalView,
            iconLabel('table', tr.normalView),
            () => instance && setWorkbookView(instance.store, 'normal'),
            active.workbookView === 'normal',
            ' demo__rb--wide',
          ),
          tool(
            'viewPageLayout',
            tr.pageLayoutView,
            iconLabel('page', tr.pageLayoutView),
            () => instance && setWorkbookView(instance.store, 'pageLayout'),
            active.workbookView === 'pageLayout',
            ' demo__rb--wide',
          ),
          tool(
            'viewPageBreakPreview',
            tr.pageBreakPreview,
            iconLabel('table', tr.pageBreakPreview),
            () => instance && setWorkbookView(instance.store, 'pageBreakPreview'),
            active.workbookView === 'pageBreakPreview',
            ' demo__rb--wide',
          ),
          watchViewMenu,
        ],
        'tiles',
      ),
      group(
        strings.viewToolbar.views,
        [
          optionSelect(
            'sheetViewSelect',
            strings.viewToolbar.views,
            instance?.store.getState().sheetViews.activeViewId ?? 'current',
            sheetViewOptions(),
            (value) => {
              if (!instance) return;
              if (value === 'current') {
                instance.store.setState((state) => ({
                  ...state,
                  sheetViews: { ...state.sheetViews, activeViewId: null },
                }));
                return;
              }
              activateSheetView(instance.store, value);
            },
            ' demo__rb-select--border',
          ),
          tool(
            'sheetViewSave',
            strings.viewToolbar.saveView,
            iconLabel('options', strings.viewToolbar.saveView),
            saveSheetViewFromRibbon,
          ),
          tool(
            'sheetViewDelete',
            strings.viewToolbar.deleteView,
            iconLabel('clear', strings.viewToolbar.deleteView),
            deleteActiveSheetViewFromRibbon,
          ),
          tool(
            'workbookObjectsView',
            strings.viewToolbar.objects,
            iconLabel('options', strings.viewToolbar.objects),
            () => instance?.openWorkbookObjects(),
          ),
          tool(
            'pivotFieldListView',
            strings.workbookObjects.pivotFieldList,
            iconLabel('options', strings.workbookObjects.pivotFieldList),
            () => {
              if (!instance?.openActivePivotFieldList()) instance?.openWorkbookObjects();
            },
          ),
        ],
        'tiles',
      ),
      group(
        tr.show,
        [
          tool(
            'viewGridlines',
            viewText.gridlines,
            iconLabel('table', viewText.gridlines),
            () => toggleViewFlag('gridlines'),
            active.gridlinesVisible,
            ' demo__rb--wide',
          ),
          tool(
            'viewHeadings',
            viewText.headings,
            iconLabel('table', viewText.headings),
            () => toggleViewFlag('headings'),
            active.headingsVisible,
            ' demo__rb--wide',
          ),
          tool(
            'viewFormulas',
            viewText.formulas,
            iconLabel('function', viewText.formulas),
            () => toggleViewFlag('formulas'),
            active.formulasVisible,
            ' demo__rb--wide',
          ),
          tool(
            'viewFormulaBar',
            viewText.formulaBar,
            iconLabel('function', viewText.formulaBar),
            onToggleFormulaBar,
            formulaBarVisible,
            ' demo__rb--wide',
          ),
          tool(
            'viewR1C1',
            'R1C1',
            iconLabel('options', 'R1C1'),
            () => toggleViewFlag('r1c1'),
            active.r1c1,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
      group(tr.window, [freezeMenu, windowMenu], 'tiles'),
      group(
        tr.zoom,
        [
          tool(
            'zoomDialog',
            `${tr.zoom}...`,
            iconLabel('zoom', tr.zoom),
            onZoomDialog,
            false,
            ' demo__rb--wide',
          ),
          tool(
            'zoomSelection',
            tr.zoomSelection,
            iconLabel('zoom', tr.zoomSelection),
            onZoomSelection,
            false,
            ' demo__rb--wide',
          ),
          tool(
            'zoom75',
            `${tr.zoom} 75%`,
            iconLabel('zoom', '75%'),
            () => onZoom(0.75),
            active.zoom === 0.75,
            ' demo__rb--mono',
          ),
          tool(
            'zoom100',
            `${tr.zoom} 100%`,
            iconLabel('zoom', '100%'),
            () => onZoom(1),
            active.zoom === 1,
            ' demo__rb--mono',
          ),
          tool(
            'zoom125',
            `${tr.zoom} 125%`,
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
            active.protected ? tr.unprotect : tr.protect,
            iconLabel('protect', active.protected ? tr.unprotect : tr.protect),
            () => instance?.toggleSheetProtection(),
            active.protected,
            ' demo__rb--wide',
          ),
        ],
        'tiles',
      ),
    ],
    help: [
      group(
        tr.tabs.help,
        [
          tool(
            'helpSearch',
            tr.tabs.help,
            iconLabel('options', tr.tabs.help),
            () => undefined,
            false,
            ' demo__rb--wide',
            true,
          ),
        ],
        'tiles',
      ),
    ],
    ...buildAddInRibbonGroups({
      group,
      iconLabel,
      addInMenu,
      pdfMenu,
      onRunScript,
      onRecordActions,
      onAllScripts,
      strings,
      tool,
      tr,
    }),
  };
};
