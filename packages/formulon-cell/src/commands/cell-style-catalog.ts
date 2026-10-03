/** Built-in cell style presets: ids, format payloads, builtin metadata, gallery groups, and platform overrides. */
import type { CellFormat } from '../store/store.js';

/** Built-in named cell styles. Each style is a partial CellFormat that
 *  `applyCellStyle` merges into the active range via `setRangeFormat`. The
 *  IDs mirror the "Cell Styles" gallery. */
export type CellStyleId =
  | 'normal'
  | 'title'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'heading4'
  | 'good'
  | 'bad'
  | 'neutral'
  | 'note'
  | 'warning'
  | 'checkCell'
  | 'explanatoryText'
  | 'inputCell'
  | 'outputCell'
  | 'calculation'
  | 'linkedCell'
  | 'totalCell'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'accent1_20'
  | 'accent2_20'
  | 'accent3_20'
  | 'accent4_20'
  | 'accent5_20'
  | 'accent6_20'
  | 'accent1_40'
  | 'accent2_40'
  | 'accent3_40'
  | 'accent4_40'
  | 'accent5_40'
  | 'accent6_40'
  | 'accent1_60'
  | 'accent2_60'
  | 'accent3_60'
  | 'accent4_60'
  | 'accent5_60'
  | 'accent6_60'
  | 'currency'
  | 'currency0'
  | 'percent'
  | 'comma'
  | 'comma0';

export interface CellStyleDef {
  id: CellStyleId;
  /** Default English label — chrome wires its own translated label and
   *  passes the id back to `applyCellStyle`. */
  label: string;
  format: Partial<CellFormat>;
  builtinId?: number;
  includedGroups?: readonly CellStyleFormatGroup[];
}

export type CellStyleFormatGroup =
  | 'number'
  | 'alignment'
  | 'font'
  | 'border'
  | 'fill'
  | 'protection';

export type CellStyleGroupId =
  | 'goodBadNeutral'
  | 'dataAndModel'
  | 'titlesAndHeadings'
  | 'themedCellStyles'
  | 'numberFormat';

export interface CellStyleGroupDef {
  id: CellStyleGroupId;
  styleIds: readonly CellStyleId[];
}

/** Spreadsheet-flavored named cell style presets. The format payloads stay close to
 *  desktop defaults so a workbook hopping between this UI and desktop spreadsheets feels
 *  consistent. Borders use the basic `'thin'`/`'medium'` styles; consumers
 *  can extend with their own gallery via `applyCellFormat` directly. */
const CELL_STYLE_DEFS: readonly CellStyleDef[] = [
  { id: 'normal', label: 'Normal', format: {} },
  {
    id: 'title',
    label: 'Title',
    format: { bold: false, fontSize: 18, color: '#1f4e79' },
  },
  {
    id: 'heading1',
    label: 'Heading 1',
    format: {
      bold: true,
      fontSize: 15,
      color: '#1f4e79',
      borders: { bottom: { style: 'medium', color: '#1f4e79' } },
    },
  },
  {
    id: 'heading2',
    label: 'Heading 2',
    format: {
      bold: true,
      fontSize: 13,
      color: '#1f4e79',
      borders: { bottom: { style: 'thin', color: '#1f4e79' } },
    },
  },
  {
    id: 'heading3',
    label: 'Heading 3',
    format: { bold: true, fontSize: 11, color: '#1f4e79' },
  },
  {
    id: 'heading4',
    label: 'Heading 4',
    format: { bold: true, italic: false, fontSize: 11, color: '#1f4e79' },
  },
  { id: 'good', label: 'Good', format: { color: '#006100', fill: '#c6efce' } },
  { id: 'bad', label: 'Bad', format: { color: '#9c0006', fill: '#ffc7ce' } },
  {
    id: 'neutral',
    label: 'Neutral',
    format: { color: '#9c5700', fill: '#ffeb9c' },
  },
  {
    id: 'note',
    label: 'Note',
    format: { fill: '#ffffcc', color: '#333333' },
  },
  {
    id: 'warning',
    label: 'Warning',
    format: { color: '#ff0000', italic: false },
  },
  {
    id: 'checkCell',
    label: 'Check Cell',
    format: { fill: '#a9d08e', color: '#375623', bold: true },
  },
  {
    id: 'explanatoryText',
    label: 'Explanatory Text',
    format: { color: '#7f7f7f', italic: true },
  },
  {
    id: 'inputCell',
    label: 'Input',
    format: { fill: '#ffcc99', color: '#3f3f76' },
  },
  {
    id: 'outputCell',
    label: 'Output',
    format: { bold: true, fill: '#f2f2f2', color: '#3f3f3f' },
  },
  {
    id: 'calculation',
    label: 'Calculation',
    format: { bold: true, italic: false, fill: '#f2f2f2', color: '#fa7d00' },
  },
  {
    id: 'linkedCell',
    label: 'Linked Cell',
    format: { color: '#fa7d00', italic: false },
  },
  {
    id: 'totalCell',
    label: 'Total',
    format: {
      bold: true,
      borders: {
        top: { style: 'thin' },
        bottom: { style: 'double' },
      },
    },
  },
  {
    id: 'accent1',
    label: 'Accent1',
    format: { color: '#ffffff', fill: '#4472c4' },
  },
  {
    id: 'accent2',
    label: 'Accent2',
    format: { color: '#ffffff', fill: '#ed7d31' },
  },
  {
    id: 'accent3',
    label: 'Accent3',
    format: { color: '#ffffff', fill: '#a5a5a5' },
  },
  {
    id: 'accent4',
    label: 'Accent4',
    format: { color: '#000000', fill: '#ffc000' },
  },
  {
    id: 'accent5',
    label: 'Accent5',
    format: { color: '#ffffff', fill: '#5b9bd5' },
  },
  {
    id: 'accent6',
    label: 'Accent6',
    format: { color: '#ffffff', fill: '#70ad47' },
  },
  {
    id: 'accent1_20',
    label: '20% - Accent1',
    format: { color: '#1f4e79', fill: '#d9e2f3' },
  },
  {
    id: 'accent2_20',
    label: '20% - Accent2',
    format: { color: '#833c0c', fill: '#fce4d6' },
  },
  {
    id: 'accent3_20',
    label: '20% - Accent3',
    format: { color: '#525252', fill: '#ededed' },
  },
  {
    id: 'accent4_20',
    label: '20% - Accent4',
    format: { color: '#7f6000', fill: '#fff2cc' },
  },
  {
    id: 'accent5_20',
    label: '20% - Accent5',
    format: { color: '#1f4e79', fill: '#ddebf7' },
  },
  {
    id: 'accent6_20',
    label: '20% - Accent6',
    format: { color: '#375623', fill: '#e2f0d9' },
  },
  {
    id: 'accent1_40',
    label: '40% - Accent1',
    format: { color: '#000000', fill: '#b4c7e7' },
  },
  {
    id: 'accent2_40',
    label: '40% - Accent2',
    format: { color: '#000000', fill: '#f8cbad' },
  },
  {
    id: 'accent3_40',
    label: '40% - Accent3',
    format: { color: '#000000', fill: '#dbdbdb' },
  },
  {
    id: 'accent4_40',
    label: '40% - Accent4',
    format: { color: '#000000', fill: '#ffe599' },
  },
  {
    id: 'accent5_40',
    label: '40% - Accent5',
    format: { color: '#000000', fill: '#bdd7ee' },
  },
  {
    id: 'accent6_40',
    label: '40% - Accent6',
    format: { color: '#000000', fill: '#c5e0b3' },
  },
  {
    id: 'accent1_60',
    label: '60% - Accent1',
    format: { color: '#000000', fill: '#8eaadb' },
  },
  {
    id: 'accent2_60',
    label: '60% - Accent2',
    format: { color: '#000000', fill: '#f4b183' },
  },
  {
    id: 'accent3_60',
    label: '60% - Accent3',
    format: { color: '#000000', fill: '#c9c9c9' },
  },
  {
    id: 'accent4_60',
    label: '60% - Accent4',
    format: { color: '#000000', fill: '#ffd966' },
  },
  {
    id: 'accent5_60',
    label: '60% - Accent5',
    format: { color: '#000000', fill: '#9dc3e6' },
  },
  {
    id: 'accent6_60',
    label: '60% - Accent6',
    format: { color: '#000000', fill: '#a8d08d' },
  },
  {
    id: 'currency',
    label: 'Currency',
    format: { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } },
  },
  {
    id: 'currency0',
    label: 'Currency [0]',
    format: { numFmt: { kind: 'currency', decimals: 0, symbol: '$' } },
  },
  {
    id: 'percent',
    label: 'Percent',
    format: { numFmt: { kind: 'percent', decimals: 0 } },
  },
  {
    id: 'comma',
    label: 'Comma',
    format: { numFmt: { kind: 'fixed', decimals: 2, thousands: true } },
  },
  {
    id: 'comma0',
    label: 'Comma [0]',
    format: { numFmt: { kind: 'fixed', decimals: 0, thousands: true } },
  },
];

const CELL_STYLE_METADATA: Partial<
  Record<CellStyleId, Pick<CellStyleDef, 'builtinId' | 'includedGroups'>>
> = {
  normal: {
    builtinId: 0,
    includedGroups: ['number', 'alignment', 'font', 'border', 'fill', 'protection'],
  },
  comma: { builtinId: 3, includedGroups: ['number'] },
  currency: { builtinId: 4, includedGroups: ['number'] },
  percent: { builtinId: 5, includedGroups: ['number'] },
  comma0: { builtinId: 6, includedGroups: ['number'] },
  currency0: { builtinId: 7, includedGroups: ['number'] },
  note: { builtinId: 10, includedGroups: ['border', 'fill'] },
  warning: { builtinId: 11, includedGroups: ['font'] },
  explanatoryText: { builtinId: 53, includedGroups: ['font'] },
  title: { builtinId: 15, includedGroups: ['font'] },
  heading1: { builtinId: 16, includedGroups: ['font', 'border'] },
  heading2: { builtinId: 17, includedGroups: ['font', 'border'] },
  heading3: { builtinId: 18, includedGroups: ['font', 'border'] },
  heading4: { builtinId: 19, includedGroups: ['font'] },
  inputCell: { builtinId: 20, includedGroups: ['font', 'border', 'fill'] },
  outputCell: { builtinId: 21, includedGroups: ['font', 'border', 'fill'] },
  calculation: { builtinId: 22, includedGroups: ['font', 'border', 'fill'] },
  checkCell: { builtinId: 23, includedGroups: ['font', 'border', 'fill'] },
  linkedCell: { builtinId: 24, includedGroups: ['font', 'border'] },
  totalCell: { builtinId: 25, includedGroups: ['font', 'border'] },
  good: { builtinId: 26, includedGroups: ['font', 'fill'] },
  bad: { builtinId: 27, includedGroups: ['font', 'fill'] },
  neutral: { builtinId: 28, includedGroups: ['font', 'fill'] },
  accent1: { builtinId: 29, includedGroups: ['font', 'fill'] },
  accent1_20: { builtinId: 30, includedGroups: ['font', 'fill'] },
  accent2: { builtinId: 33, includedGroups: ['font', 'fill'] },
  accent2_20: { builtinId: 34, includedGroups: ['font', 'fill'] },
  accent3: { builtinId: 37, includedGroups: ['font', 'fill'] },
  accent3_20: { builtinId: 38, includedGroups: ['font', 'fill'] },
  accent4: { builtinId: 41, includedGroups: ['font', 'fill'] },
  accent4_20: { builtinId: 42, includedGroups: ['font', 'fill'] },
  accent5: { builtinId: 45, includedGroups: ['font', 'fill'] },
  accent5_20: { builtinId: 46, includedGroups: ['font', 'fill'] },
  accent6: { builtinId: 49, includedGroups: ['font', 'fill'] },
  accent6_20: { builtinId: 50, includedGroups: ['font', 'fill'] },
  accent1_40: { builtinId: 31, includedGroups: ['font', 'fill'] },
  accent1_60: { builtinId: 32, includedGroups: ['font', 'fill'] },
  accent2_40: { builtinId: 35, includedGroups: ['font', 'fill'] },
  accent2_60: { builtinId: 36, includedGroups: ['font', 'fill'] },
  accent3_40: { builtinId: 39, includedGroups: ['font', 'fill'] },
  accent3_60: { builtinId: 40, includedGroups: ['font', 'fill'] },
  accent4_40: { builtinId: 43, includedGroups: ['font', 'fill'] },
  accent4_60: { builtinId: 44, includedGroups: ['font', 'fill'] },
  accent5_40: { builtinId: 47, includedGroups: ['font', 'fill'] },
  accent5_60: { builtinId: 48, includedGroups: ['font', 'fill'] },
  accent6_40: { builtinId: 51, includedGroups: ['font', 'fill'] },
  accent6_60: { builtinId: 52, includedGroups: ['font', 'fill'] },
};

export const CELL_STYLES: readonly CellStyleDef[] = CELL_STYLE_DEFS.map((style) => ({
  ...style,
  ...CELL_STYLE_METADATA[style.id],
}));

export const CELL_STYLE_GROUPS: readonly CellStyleGroupDef[] = [
  {
    id: 'goodBadNeutral',
    styleIds: ['normal', 'good', 'bad', 'neutral'],
  },
  {
    id: 'dataAndModel',
    styleIds: [
      'note',
      'warning',
      'checkCell',
      'explanatoryText',
      'inputCell',
      'outputCell',
      'calculation',
      'linkedCell',
      'totalCell',
    ],
  },
  {
    id: 'titlesAndHeadings',
    styleIds: ['title', 'heading1', 'heading2', 'heading3', 'heading4'],
  },
  {
    id: 'themedCellStyles',
    styleIds: [
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
      'accent1_20',
      'accent2_20',
      'accent3_20',
      'accent4_20',
      'accent5_20',
      'accent6_20',
      'accent1_40',
      'accent2_40',
      'accent3_40',
      'accent4_40',
      'accent5_40',
      'accent6_40',
      'accent1_60',
      'accent2_60',
      'accent3_60',
      'accent4_60',
      'accent5_60',
      'accent6_60',
    ],
  },
  {
    id: 'numberFormat',
    styleIds: ['currency', 'currency0', 'percent', 'comma', 'comma0'],
  },
];

export const EXCEL365_MAC_STYLE_OVERRIDES: Partial<Record<CellStyleId, Partial<CellFormat>>> = {
  title: { color: '#0e2841' },
  heading1: { color: '#0e2841' },
  heading2: { color: '#0e2841' },
  heading3: { color: '#0e2841' },
  heading4: { color: '#0e2841' },
  checkCell: { color: '#ffffff', fill: '#a5a5a5' },
  accent1: { color: '#ffffff', fill: '#156082' },
  accent1_20: { color: '#000000', fill: '#c0e6f5' },
  accent1_40: { color: '#000000', fill: '#83cceb' },
  accent1_60: { color: '#000000', fill: '#44b3e1' },
  accent2: { color: '#ffffff', fill: '#e97132' },
  accent2_20: { color: '#000000', fill: '#fbe2d5' },
  accent2_40: { color: '#000000', fill: '#f7c7ac' },
  accent2_60: { color: '#000000', fill: '#f1a983' },
  accent3: { color: '#ffffff', fill: '#196b24' },
  accent3_20: { color: '#000000', fill: '#c1f0c8' },
  accent3_40: { color: '#000000', fill: '#83e28e' },
  accent3_60: { color: '#000000', fill: '#47d359' },
  accent4: { color: '#ffffff', fill: '#0f9ed5' },
  accent4_20: { color: '#000000', fill: '#caedfb' },
  accent4_40: { color: '#000000', fill: '#94dcf8' },
  accent4_60: { color: '#000000', fill: '#61cbf3' },
  accent5: { color: '#ffffff', fill: '#a02b93' },
  accent5_20: { color: '#000000', fill: '#f2ceef' },
  accent5_40: { color: '#000000', fill: '#e49edd' },
  accent5_60: { color: '#000000', fill: '#d86dcd' },
  accent6: { color: '#ffffff', fill: '#4ea72e' },
  accent6_20: { color: '#000000', fill: '#daf2d0' },
  accent6_40: { color: '#000000', fill: '#b5e6a2' },
  accent6_60: { color: '#000000', fill: '#8ed973' },
};
