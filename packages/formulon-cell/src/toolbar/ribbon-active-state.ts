import type { BorderPreset as CommandBorderPreset } from '../commands/border-presets.js';
import { activeCellStyleId } from '../commands/cell-styles.js';
import { commentAt, noteVisibleAt } from '../commands/comment.js';
import { activeFilterRange } from '../commands/filter-ranges.js';
import { tableForCell } from '../commands/format-as-table.js';
import { mergeAt } from '../commands/merge.js';
import type { MarginPreset } from '../commands/page-setup.js';
import { marginPresetOf, pageSetupForSheet } from '../commands/page-setup.js';
import { hiddenExtentInSelection } from '../commands/row-col-layout.js';
import { effectiveCellFormat } from '../engine/cell-format-projection.js';
import { fontCssPxToDisplayPoints, fontPointsToCssPx } from '../engine/font-units.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetInstance } from '../mount.js';
import { outlineSummaryFor } from '../store/outline-summary.js';
import { rangesIntersect } from '../store/selection-geometry.js';
import type {
  CellBorderStyle,
  NumFmt,
  PageOrientation,
  PaperSize,
  WorkbookViewMode,
} from '../store/types.js';
import { projectDisabledState, projectValueName, RIBBON_DISABLED_REASON_KEY } from './menu-a11y.js';
import { projectPressedState } from './ribbon/button.js';
import { RIBBON_VIEW_MODES } from './ribbon/command-tables.js';
import { projectSparklineDesign } from './ribbon/sparkline-design.js';
import { projectPivotDesign, projectTableDesign } from './ribbon/table-design.js';

/**
 * Ribbon active-state contract shared by every framework wrapper.
 *
 * Both `@libraz/formulon-cell-react` and `@libraz/formulon-cell-vue`
 * derive their ribbon button states (bold pressed?, current font, page
 * orientation, …) from the same projection over the spreadsheet store,
 * so the canonical shape, defaults, and projection live here in core.
 * Wrappers re-export these so consumers keep one source of truth even
 * when mixing frameworks in the same app.
 */
export interface ActiveState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  showPhonetic: boolean;
  strike: boolean;
  alignLeft: boolean;
  alignCenter: boolean;
  alignRight: boolean;
  vAlignTop: boolean;
  vAlignMiddle: boolean;
  vAlignBottom: boolean;
  wrapText: boolean;
  merged: boolean;
  mergeCenter: boolean;
  conditionalFormatting: boolean;
  formatAsTable: boolean;
  cellStyle: string | null;
  textOrientation:
    | 'angleCounterclockwise'
    | 'angleClockwise'
    | 'verticalText'
    | 'rotateTextUp'
    | 'rotateTextDown'
    | 'horizontalText';
  currency: boolean;
  percent: boolean;
  commaStyle: boolean;
  numberFormat: string;
  frozen: boolean;
  filterOn: boolean;
  rowsHidden: boolean;
  colsHidden: boolean;
  protected: boolean;
  zoom: number;
  formatPainterArmed: boolean;
  hasComment: boolean;
  noteVisible: boolean;
  allNotesVisible: boolean;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
  pageOrientation: PageOrientation;
  paperSize: PaperSize;
  pageScale: number;
  fitWidth: number | null;
  fitHeight: number | null;
  gridlinesVisible: boolean;
  headingsVisible: boolean;
  printGridlines: boolean;
  printHeadings: boolean;
  formulasVisible: boolean;
  outlineSummaryBelow: boolean;
  outlineSummaryRight: boolean;
  outlineSymbolsVisible: boolean;
  splitPanes: boolean;
  zerosVisible: boolean;
  rightToLeft: boolean;
  workbookView: WorkbookViewMode;
  r1c1: boolean;
  calcMode: 0 | 1 | 2 | null;
  /** Closest named preset for the active sheet's margins, or `null`
   *  when the user has set custom values via the Page Setup dialog. */
  marginPreset: MarginPreset | null;
}

export const EMPTY_ACTIVE_STATE: ActiveState = {
  bold: false,
  italic: false,
  underline: false,
  showPhonetic: false,
  strike: false,
  alignLeft: false,
  alignCenter: false,
  alignRight: false,
  vAlignTop: false,
  vAlignMiddle: false,
  vAlignBottom: true,
  wrapText: false,
  merged: false,
  mergeCenter: false,
  conditionalFormatting: false,
  formatAsTable: false,
  cellStyle: null,
  textOrientation: 'horizontalText',
  currency: false,
  percent: false,
  commaStyle: false,
  numberFormat: 'general',
  frozen: false,
  filterOn: false,
  rowsHidden: false,
  colsHidden: false,
  protected: false,
  zoom: 1,
  formatPainterArmed: false,
  hasComment: false,
  noteVisible: false,
  allNotesVisible: false,
  fontFamily: 'Aptos',
  fontSize: 11,
  fontColor: '#201f1e',
  fillColor: '#ffffff',
  pageOrientation: 'portrait',
  paperSize: 'A4',
  pageScale: 1,
  fitWidth: null,
  fitHeight: null,
  gridlinesVisible: true,
  headingsVisible: true,
  printGridlines: false,
  printHeadings: false,
  formulasVisible: false,
  outlineSummaryBelow: true,
  outlineSummaryRight: true,
  outlineSymbolsVisible: true,
  splitPanes: false,
  zerosVisible: true,
  rightToLeft: false,
  workbookView: 'normal',
  r1c1: false,
  calcMode: null,
  marginPreset: 'normal',
};

/** Groups every ribbon button under `host` by its `data-ribbon-command`, in
 *  one DOM pass, so a projection over many commands does not rescan the host
 *  per command. */
export const indexRibbonButtons = (host: HTMLElement): Map<string, HTMLButtonElement[]> => {
  const index = new Map<string, HTMLButtonElement[]>();
  for (const button of host.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')) {
    const command = button.dataset.ribbonCommand;
    if (command === undefined) continue;
    const group = index.get(command);
    if (group) group.push(button);
    else index.set(command, [button]);
  }
  return index;
};

/** Labels the toolbar owner rendered; the projection never infers a locale of its own. */
export interface RibbonActiveStateText {
  background?: string;
  sheetBackgroundClear?: string;
}

export const RIBBON_ACTIVE_COMMANDS: ReadonlyMap<string, keyof ActiveState> = new Map([
  ['bold', 'bold'],
  ['italic', 'italic'],
  ['underline', 'underline'],
  ['showPhonetic', 'showPhonetic'],
  ['strike', 'strike'],
  ['alignL', 'alignLeft'],
  ['alignC', 'alignCenter'],
  ['alignR', 'alignRight'],
  ['top', 'vAlignTop'],
  ['middle', 'vAlignMiddle'],
  ['bottomAlign', 'vAlignBottom'],
  ['wrap', 'wrapText'],
  ['merge', 'merged'],
  ['currency', 'currency'],
  ['percent', 'percent'],
  ['comma', 'commaStyle'],
  ['conditional', 'conditionalFormatting'],
  ['formatTableHome', 'formatAsTable'],
  ['freeze', 'frozen'],
  ['filter', 'filterOn'],
  ['protect', 'protected'],
  ['protectReview', 'protected'],
  ['viewGridlines', 'gridlinesVisible'],
  ['viewHeadings', 'headingsVisible'],
  ['viewFormulas', 'formulasVisible'],
  ['outlineSummaryBelow', 'outlineSummaryBelow'],
  ['outlineSummaryRight', 'outlineSummaryRight'],
  ['outlineSymbols', 'outlineSymbolsVisible'],
  ['split', 'splitPanes'],
  ['viewZeros', 'zerosVisible'],
  ['viewRightToLeft', 'rightToLeft'],
  ['viewR1C1', 'r1c1'],
  ['viewNormal', 'workbookView'],
  ['viewPageLayout', 'workbookView'],
  ['viewPageBreakPreview', 'workbookView'],
  ['formatPainter', 'formatPainterArmed'],
]);

const numberFormatOf = (fmt: NumFmt | undefined): string => {
  if (!fmt) return 'general';
  switch (fmt.kind) {
    case 'fixed':
      return 'fixed';
    case 'currency':
      return 'currency';
    case 'accounting':
      return 'accounting';
    case 'date':
      return fmt.pattern.includes('mmmm') || fmt.pattern.includes('"年"')
        ? 'longDate'
        : 'shortDate';
    case 'time':
      return 'time';
    case 'percent':
      return 'percent';
    case 'custom':
      return fmt.pattern.includes('?/?') ? 'fraction' : 'general';
    case 'scientific':
      return 'scientific';
    case 'text':
      return 'text';
    case 'datetime':
      return 'shortDate';
    default:
      return 'general';
  }
};

type BorderPresetLabelKey =
  | 'noBorder'
  | 'outsideBorders'
  | 'thickOutsideBorders'
  | 'allBorders'
  | 'insideBorders'
  | 'insideHorizontalBorder'
  | 'insideVerticalBorder'
  | 'topBorder'
  | 'bottomBorder'
  | 'leftBorder'
  | 'rightBorder'
  | 'doubleBottomBorder'
  | 'thickBottomBorder'
  | 'topAndBottomBorder'
  | 'topAndThickBottomBorder'
  | 'topAndDoubleBottomBorder'
  | 'diagonalDownBorder'
  | 'diagonalUpBorder';

type BorderStyleLabelKey =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export const BORDER_PRESETS: {
  value: CommandBorderPreset;
  label: string;
  labelKey: BorderPresetLabelKey;
}[] = [
  { value: 'none', label: 'No Border', labelKey: 'noBorder' },
  { value: 'outline', label: 'Outside Borders', labelKey: 'outsideBorders' },
  { value: 'thickOutline', label: 'Thick Outside Borders', labelKey: 'thickOutsideBorders' },
  { value: 'all', label: 'All Borders', labelKey: 'allBorders' },
  { value: 'inside', label: 'Inside Borders', labelKey: 'insideBorders' },
  {
    value: 'insideHorizontal',
    label: 'Inside Horizontal Border',
    labelKey: 'insideHorizontalBorder',
  },
  { value: 'insideVertical', label: 'Inside Vertical Border', labelKey: 'insideVerticalBorder' },
  { value: 'top', label: 'Top Border', labelKey: 'topBorder' },
  { value: 'bottom', label: 'Bottom Border', labelKey: 'bottomBorder' },
  { value: 'left', label: 'Left Border', labelKey: 'leftBorder' },
  { value: 'right', label: 'Right Border', labelKey: 'rightBorder' },
  { value: 'doubleBottom', label: 'Double Bottom', labelKey: 'doubleBottomBorder' },
  { value: 'thickBottom', label: 'Thick Bottom Border', labelKey: 'thickBottomBorder' },
  { value: 'topAndBottom', label: 'Top and Bottom Border', labelKey: 'topAndBottomBorder' },
  {
    value: 'topAndThickBottom',
    label: 'Top and Thick Bottom Border',
    labelKey: 'topAndThickBottomBorder',
  },
  {
    value: 'topAndDoubleBottom',
    label: 'Top and Bottom Double Border',
    labelKey: 'topAndDoubleBottomBorder',
  },
  { value: 'diagonalDown', label: 'Diagonal Down Border', labelKey: 'diagonalDownBorder' },
  { value: 'diagonalUp', label: 'Diagonal Up Border', labelKey: 'diagonalUpBorder' },
];

export const BORDER_STYLES: {
  value: CellBorderStyle;
  label: string;
  labelKey: BorderStyleLabelKey;
}[] = [
  { value: 'thin', label: 'Thin', labelKey: 'thin' },
  { value: 'medium', label: 'Medium', labelKey: 'medium' },
  { value: 'thick', label: 'Thick', labelKey: 'thick' },
  { value: 'dashed', label: 'Dashed', labelKey: 'dashed' },
  { value: 'dotted', label: 'Dotted', labelKey: 'dotted' },
  { value: 'double', label: 'Double', labelKey: 'double' },
  { value: 'hair', label: 'Hairline', labelKey: 'hair' },
  { value: 'mediumDashed', label: 'Medium Dashed', labelKey: 'mediumDashed' },
  { value: 'dashDot', label: 'Dash Dot', labelKey: 'dashDot' },
  { value: 'mediumDashDot', label: 'Medium Dash Dot', labelKey: 'mediumDashDot' },
  { value: 'dashDotDot', label: 'Dash Dot Dot', labelKey: 'dashDotDot' },
  { value: 'mediumDashDotDot', label: 'Medium Dash Dot Dot', labelKey: 'mediumDashDotDot' },
  { value: 'slantDashDot', label: 'Slant Dash Dot', labelKey: 'slantDashDot' },
];

export const localizeBorderPresets = (
  ribbon: Strings['ribbon'],
): { value: CommandBorderPreset; label: string }[] =>
  BORDER_PRESETS.map((preset) => ({ value: preset.value, label: ribbon[preset.labelKey] }));

export const localizeBorderStyles = (
  ribbon: Strings['ribbon'],
): { value: CellBorderStyle; label: string }[] =>
  BORDER_STYLES.map((style) => ({ value: style.value, label: ribbon[style.labelKey] }));

export const projectActiveState = (inst: SpreadsheetInstance): ActiveState => {
  const s = inst.store.getState();
  const a = s.selection.active;
  const r = s.selection.range;
  const activeMerge = mergeAt(s, a);
  const formatAddr = activeMerge
    ? { sheet: activeMerge.sheet, row: activeMerge.r0, col: activeMerge.c0 }
    : a;
  const f = effectiveCellFormat(inst.workbook, s, formatAddr);
  const displayFont = inst.workbook.workbookDisplayFont;
  const displayFamily = displayFont?.name ?? 'Aptos';
  const displaySize = displayFont ? fontPointsToCssPx(displayFont.size) : 11;
  const hasConditionalFormatting = s.conditional.rules.some((rule) =>
    rangesIntersect(r, rule.range),
  );
  const activeTable = tableForCell(s.tables.tables, a.sheet, a.row, a.col);
  const setup = pageSetupForSheet(s, s.data.sheetIndex);
  const outlineSummary = outlineSummaryFor(s.layout, s.data.sheetIndex);
  return {
    bold: !!f?.bold,
    italic: !!f?.italic,
    underline: !!f?.underline,
    showPhonetic: f?.showPhonetic === true,
    strike: !!f?.strike,
    alignLeft: f?.align === 'left',
    alignCenter: f?.align === 'center',
    alignRight: f?.align === 'right',
    vAlignTop: f?.vAlign === 'top',
    vAlignMiddle: f?.vAlign === 'middle',
    vAlignBottom: f?.vAlign == null || f.vAlign === 'bottom',
    wrapText: f?.wrap === true,
    merged: activeMerge != null,
    mergeCenter: activeMerge != null && f?.align === 'center',
    conditionalFormatting: hasConditionalFormatting,
    formatAsTable: activeTable != null,
    cellStyle: activeCellStyleId(s),
    textOrientation:
      f?.rotation === 'vertical'
        ? 'verticalText'
        : f?.rotation === 45
          ? 'angleCounterclockwise'
          : f?.rotation === -45
            ? 'angleClockwise'
            : f?.rotation === 90
              ? 'rotateTextUp'
              : f?.rotation === -90
                ? 'rotateTextDown'
                : 'horizontalText',
    currency: f?.numFmt?.kind === 'currency',
    percent: f?.numFmt?.kind === 'percent',
    commaStyle: f?.numFmt?.kind === 'fixed' && f.numFmt.thousands === true,
    numberFormat: numberFormatOf(f?.numFmt),
    frozen: s.layout.splitPanes !== true && (s.layout.freezeRows > 0 || s.layout.freezeCols > 0),
    filterOn: activeFilterRange(s) !== null,
    rowsHidden: hiddenExtentInSelection(s.layout, 'row', r.r0, r.r1) !== null,
    colsHidden: hiddenExtentInSelection(s.layout, 'col', r.c0, r.c1) !== null,
    protected: inst.isSheetProtected(),
    zoom: s.viewport.zoom,
    formatPainterArmed: !!inst.formatPainter?.isActive(),
    hasComment: commentAt(s, a) != null,
    noteVisible: noteVisibleAt(s, a),
    allNotesVisible: s.ui.showAllNotes,
    fontFamily: f?.fontFamily ?? displayFamily,
    // CellFormat stores CSS pixels for paint; Excel's ribbon exposes points.
    fontSize: fontCssPxToDisplayPoints(f?.fontSize ?? displaySize),
    fontColor: f?.color ?? '#201f1e',
    fillColor: f?.fill ?? '#ffffff',
    pageOrientation: setup.orientation,
    paperSize: setup.paperSize,
    pageScale: setup.scale ?? 1,
    fitWidth: setup.fitWidth ?? null,
    fitHeight: setup.fitHeight ?? null,
    gridlinesVisible: s.ui.showGridLines !== false,
    headingsVisible: s.ui.showHeaders !== false,
    printGridlines: setup.showGridlines === true,
    printHeadings: setup.showHeadings === true,
    formulasVisible: !!s.ui.showFormulas,
    outlineSummaryBelow: outlineSummary.below,
    outlineSummaryRight: outlineSummary.right,
    outlineSymbolsVisible: s.ui.showOutlineSymbols !== false,
    splitPanes: s.layout.splitPanes === true,
    zerosVisible: s.ui.showZeros !== false,
    rightToLeft: s.ui.rightToLeft === true,
    workbookView: s.ui.workbookView,
    r1c1: !!s.ui.r1c1,
    calcMode: inst.workbook.calcMode(),
    marginPreset: marginPresetOf(setup.margins),
  };
};

/** The Draw tab's Pen and Eraser set the host's ink mode; with no drawing hook they would do
 *  nothing, so they stay greyed. Re-enables only what this projection greyed. */
export const projectDrawingCommands = (host: HTMLElement, available: boolean): void => {
  for (const command of ['drawPen', 'drawErase']) {
    for (const button of host.querySelectorAll<HTMLButtonElement>(
      `[data-ribbon-command="${command}"]`,
    )) {
      if (available && button.dataset.drawingHookMissing !== 'true') continue;
      projectDisabledState(button, !available, null);
      if (available) delete button.dataset.drawingHookMissing;
      else button.dataset.drawingHookMissing = 'true';
    }
  }
};

/** Projects the active state onto the ribbon buttons under `host`. */
export const projectRibbonActiveState = (
  host: HTMLElement,
  instance: SpreadsheetInstance | null,
  text: Required<RibbonActiveStateText>,
): void => {
  if (!instance) return;
  const active = projectActiveState(instance);
  const EventCtor = host.ownerDocument.defaultView?.Event ?? Event;
  for (const control of host.querySelectorAll<HTMLElement>('[data-ribbon-select]')) {
    control.dispatchEvent(new EventCtor('fc-ribbon-control-refresh'));

    // The font family field intentionally truncates inside its Excel-sized
    // input. Keep the complete rendered family available to assistive tech
    // and the native hover tooltip even when the visible value is ellipsized.
    if (control.dataset.ribbonSelect === 'fontFamily') {
      const button = control.querySelector<HTMLButtonElement>('.fc-tb__rb-dd__btn');
      const value = control.querySelector<HTMLElement>('.fc-tb__rb-dd__value');
      const renderedValue = value?.textContent?.trim();
      if (button && renderedValue) {
        const fullLabel = projectValueName(button, renderedValue);
        button.setAttribute('title', fullLabel);
        value?.setAttribute('title', renderedValue);
      }
    }
  }
  projectTableDesign(host, instance);
  projectPivotDesign(host, instance);
  projectSparklineDesign(host, instance);
  const buttons = indexRibbonButtons(host);
  // Authoring a split the engine cannot save is refused up front.
  const splitSaved = instance.workbook?.capabilities?.sheetSplit === true;
  for (const button of buttons.get('split') ?? []) {
    // Re-enable only what this projection disabled; an embedding policy owns the rest.
    if (splitSaved && button.dataset.splitNotSaved !== 'true') continue;
    if (splitSaved) delete button.dataset.splitNotSaved;
    else button.dataset.splitNotSaved = 'true';
    projectDisabledState(
      button,
      !splitSaved,
      splitSaved ? null : instance.i18n.strings.ribbonMenu.splitNotSaved,
      { datasetKey: RIBBON_DISABLED_REASON_KEY },
    );
  }
  for (const [command, key] of RIBBON_ACTIVE_COMMANDS) {
    const button = buttons.get(command)?.[0];
    if (!button) continue;
    const viewMode = RIBBON_VIEW_MODES[command];
    projectPressedState(button, viewMode ? active.workbookView === viewMode : Boolean(active[key]));
  }
  const fieldListOpen = instance.store.getState().ui.pivotFieldListOpen;
  for (const command of ['pivotFieldList', 'pivotFieldListView']) {
    for (const button of buttons.get(command) ?? []) projectPressedState(button, fieldListOpen);
  }

  const sheetBackground = host.querySelector<HTMLButtonElement>(
    '[data-ribbon-command="sheetBackground"]',
  );
  if (sheetBackground) {
    const state = instance.store.getState();
    const hasBackground = state.ui.sheetBackgroundImages.has(state.data.sheetIndex);
    const renderedLabel =
      sheetBackground.querySelector('span')?.textContent?.trim() ||
      sheetBackground.getAttribute('aria-label')?.trim() ||
      sheetBackground.title.trim();
    const baseLabel = text.background || renderedLabel;
    const resolvedClearLabel = text.sheetBackgroundClear;
    const label = hasBackground ? resolvedClearLabel : baseLabel;
    sheetBackground.title = label;
    sheetBackground.setAttribute('aria-label', label);
    const labelEl = sheetBackground.querySelector('span');
    if (labelEl) labelEl.textContent = label;
  }
};
