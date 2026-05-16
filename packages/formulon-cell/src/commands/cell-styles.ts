import type { Range } from '../engine/types.js';
import { type CellFormat, mutators, type SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { recordFormatChange } from './history.js';

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
}

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
export const CELL_STYLES: readonly CellStyleDef[] = [
  { id: 'normal', label: 'Normal', format: {} },
  {
    id: 'title',
    label: 'Title',
    format: { bold: true, fontSize: 18, color: '#1f4e79' },
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
    format: { bold: true, color: '#1f4e79' },
  },
  {
    id: 'heading4',
    label: 'Heading 4',
    format: { italic: true, color: '#1f4e79' },
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
    format: { color: '#ff0000', italic: true },
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
    format: { bold: true, italic: true, fill: '#f2f2f2', color: '#fa7d00' },
  },
  {
    id: 'linkedCell',
    label: 'Linked Cell',
    format: { color: '#fa7d00', italic: true },
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
    ],
  },
  {
    id: 'numberFormat',
    styleIds: ['currency', 'currency0', 'percent', 'comma', 'comma0'],
  },
];

const STYLE_BY_ID = new Map<CellStyleId, CellStyleDef>(CELL_STYLES.map((s) => [s.id, s]));

export function getCellStyle(id: CellStyleId): CellStyleDef | undefined {
  return STYLE_BY_ID.get(id);
}

/** Apply a named style to `range`. Wraps the format mutation in a single
 *  history entry so Cmd+Z reverts the whole gallery click. The `normal`
 *  style is a clear — it strips every format field instead of merging. */
export function applyCellStyle(
  store: SpreadsheetStore,
  history: History | null,
  range: Range,
  id: CellStyleId,
): void {
  const def = STYLE_BY_ID.get(id);
  if (!def) return;
  recordFormatChange(history, store, () => {
    if (id === 'normal') {
      // Clear by overwriting every format field with undefined. setRangeFormat
      //  merges with `Object.assign`, so explicit `undefined`s win — matching
      //  the spreadsheet's "Normal" reset behavior.
      mutators.setRangeFormat(store, range, {
        bold: undefined,
        italic: undefined,
        underline: undefined,
        strike: undefined,
        align: undefined,
        vAlign: undefined,
        wrap: undefined,
        indent: undefined,
        rotation: undefined,
        borders: undefined,
        color: undefined,
        fill: undefined,
        fontFamily: undefined,
        fontSize: undefined,
        numFmt: undefined,
        cellStyle: undefined,
      });
      return;
    }
    mutators.setRangeFormat(store, range, { ...def.format, cellStyle: id });
  });
}
