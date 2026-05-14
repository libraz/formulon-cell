import type {
  CellBorderStyle,
  MarginPreset,
  PageOrientation,
  PaperSize,
  SpreadsheetInstance,
} from '@libraz/formulon-cell';
import type { Dispatch, ReactElement, SetStateAction } from 'react';
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
  borderPresets: readonly { value: BorderPreset; label: string }[];
  borderStyle: CellBorderStyle;
  borderStyles: readonly { value: CellBorderStyle; label: string }[];
  color: (
    id: string,
    title: string,
    value: string,
    onChange: (value: string) => void,
    label: ReactElement,
  ) => ReactElement;
  group: GroupFn;
  iconLabel: (icon: IconName, text: string) => ReactElement;
  instance: SpreadsheetInstance | null;
  lang: 'ja' | 'en';
  onAutoSum: () => void;
  onBorderPreset: (preset: BorderPreset) => void;
  onCopy: () => void;
  onCut: () => void;
  onDeleteCols: () => void;
  onDeleteRows: () => void;
  onFilterToggle: () => void;
  onFormatAsTable: () => void;
  onFormatPainter: () => void;
  onFreezeToggle: () => void;
  onInsertCols: () => void;
  onInsertRows: () => void;
  onMarginPreset: (next: MarginPreset) => void;
  onMerge: () => void;
  onPageOrientation: (next: PageOrientation) => void;
  onPaperSize: (next: PaperSize) => void;
  onPaste: () => void;
  onRedo: () => void;
  onRemoveDuplicates: () => void;
  onSort: (direction: 'asc' | 'desc') => void;
  onToggleColsHidden: () => void;
  onToggleRowsHidden: () => void;
  onUndo: () => void;
  onZoom: (zoom: number) => void;
  optionSelect: OptionSelectFn;
  rowBreak: (id: string) => ReactElement;
  select: SelectFn;
  setBorderStyle: Dispatch<SetStateAction<CellBorderStyle>>;
  tool: ToolFn;
  tr: ToolbarText;
  wrapFormat: (
    fn: (
      state: ReturnType<SpreadsheetInstance['store']['getState']>,
      store: SpreadsheetInstance['store'],
    ) => void,
  ) => void;
}
