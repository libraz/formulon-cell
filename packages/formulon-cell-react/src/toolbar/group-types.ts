import type {
  CellBorderStyle,
  MarginPreset,
  PageOrientation,
  PaperSize,
  RibbonReportItem,
  SpreadsheetInstance,
  Strings,
} from '@libraz/formulon-cell';
import type { ReactElement } from 'react';
import type { IconName } from './icons.js';
import type { ActiveState, BorderPreset } from './model.js';
import type { ToolbarText } from './translations.js';

export type ToolFn = (
  id: string,
  title: string,
  label: string | ReactElement,
  onClick: () => void,
  isActive?: boolean,
  extra?: string,
  disabled?: boolean,
  allowWithoutInstance?: boolean,
) => ReactElement;

export type GroupFn = (title: string, children: ReactElement[], variant?: string) => ReactElement;

export type SelectFn = (
  id: string,
  title: string,
  value: string | number,
  values: readonly (string | number)[],
  onChange: (value: string) => void,
  extra?: string,
) => ReactElement;

export type OptionSelectFn = <T extends string>(
  id: string,
  title: string,
  value: T,
  options: readonly { value: T; label: string }[],
  onChange: (value: T) => void,
  extra?: string,
) => ReactElement;

export interface BuildRibbonGroupsOptions {
  active: ActiveState;
  autosumFormulaMenu: ReactElement;
  autosumMenu: ReactElement;
  borderPresets: readonly { value: BorderPreset; label: string }[];
  borderColor: string;
  borderStyle: CellBorderStyle;
  borderStyles: readonly { value: CellBorderStyle; label: string }[];
  calcOptionsMenu: ReactElement;
  addInMenu: ReactElement;
  pasteMenu: ReactElement;
  pivotTableMenu: ReactElement;
  pictureInsertMenu: ReactElement;
  shapesInsertMenu: ReactElement;
  screenshotInsertMenu: ReactElement;
  pdfMenu: ReactElement;
  protectionMenu: ReactElement;
  color: (
    id: string,
    title: string,
    value: string,
    onChange: (value: string) => void,
    label: ReactElement,
  ) => ReactElement;
  group: GroupFn;
  iconLabel: (icon: IconName, text: string) => ReactElement;
  cellInsertMenu: ReactElement;
  cellDeleteMenu: ReactElement;
  cellFormatMenu: ReactElement;
  cellStylesMenu: ReactElement;
  chartMenu: ReactElement;
  conditionalMenu: ReactElement;
  clearMenu: ReactElement;
  clearArrowsMenu: ReactElement;
  dataFilterMenu: ReactElement;
  dataSortMenu: ReactElement;
  dataValidationMenu: ReactElement;
  definedNamesMenu: ReactElement;
  definedNamesInsertMenu: ReactElement;
  deleteCommentMenu: ReactElement;
  errorCheckingMenu: ReactElement;
  freezeMenu: ReactElement;
  windowMenu: ReactElement;
  formatTableHomeMenu: ReactElement;
  formatTableInsertMenu: ReactElement;
  fillMenu: ReactElement;
  findMenu: ReactElement;
  functionDateTimeMenu: ReactElement;
  functionFinancialMenu: ReactElement;
  functionLogicalMenu: ReactElement;
  functionLookupMenu: ReactElement;
  functionMathTrigMenu: ReactElement;
  functionTextMenu: ReactElement;
  hyperlinkMenu: ReactElement;
  outlineGroupMenu: ReactElement;
  outlineUngroupMenu: ReactElement;
  printAreaMenu: ReactElement;
  pageBreaksMenu: ReactElement;
  sheetBackgroundMenu: ReactElement;
  printTitlesMenu: ReactElement;
  sortMenu: ReactElement;
  symbolMenu: ReactElement;
  themeMenu: ReactElement;
  textOrientationMenu: ReactElement;
  textToColumnsMenu: ReactElement;
  watchMenu: ReactElement;
  watchViewMenu: ReactElement;
  formulaBarVisible: boolean;
  instance: SpreadsheetInstance | null;
  lang: 'ja' | 'en';
  locale: string;
  strings: Strings;
  workbookStructureProtected: boolean;
  mergeMenu: ReactElement;
  onBorderPreset: (preset: BorderPreset) => void;
  onCopy: () => void;
  onCut: () => void;
  onDeleteCols: () => void;
  onDeleteRows: () => void;
  onAddIn?: () => void;
  onFilterToggle: () => void;
  onFormatPainter: () => void;
  onDrawEraser?: () => void;
  onDrawPen?: () => void;
  onInsertCols: () => void;
  onInsertRows: () => void;
  onMarginPreset: (next: MarginPreset) => void;
  onNumberFormat: (next: string) => void;
  onPageOrientation: (next: PageOrientation) => void;
  onPaperSize: (next: PaperSize) => void;
  onPaste: () => void;
  onProtectWorkbook?: () => void;
  onInspectWorkbook?: () => void;
  onRedo: () => void;
  onRemoveDuplicates: () => void;
  onScaleFit: (axis: 'width' | 'height', pages: string) => void;
  onScalePercent: (percent: string) => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onRecordActions?: () => void;
  onAllScripts?: () => void;
  onBuiltInReview?: (title: string, items: readonly RibbonReportItem[]) => void;
  onSort: (direction: 'asc' | 'desc') => void;
  onSpellingReview?: () => void;
  onTranslate?: () => void;
  onToggleColsHidden: () => void;
  onToggleFormulaBar: () => void;
  onToggleRowsHidden: () => void;
  onUndo: () => void;
  onZoom: (zoom: number) => void;
  onZoomDialog: () => void;
  onZoomSelection: () => void;
  optionSelect: OptionSelectFn;
  rowBreak: (id: string) => ReactElement;
  select: SelectFn;
  setBorderStyle: (value: CellBorderStyle) => void;
  setBorderColor: (value: string) => void;
  tool: ToolFn;
  tr: ToolbarText;
  wrapFormat: (
    fn: (
      state: ReturnType<SpreadsheetInstance['store']['getState']>,
      store: SpreadsheetInstance['store'],
    ) => void,
  ) => void;
}
