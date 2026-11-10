import type {
  SheetOoxml as EngineSheetOoxml,
  WorkbookOoxml as EngineWorkbookOoxml,
} from '@libraz/formulon';
import type { MemberSpec, SheetContext, WorkbookContext } from './facade-context.js';

/**
 * The OOXML-level layer of a workbook, `Workbook.ooxml`: the style record tables (cell XFs, fonts,
 * fills, borders, number formats, differential formats) and the parts kept verbatim. Members are
 * the engine's `Workbook.ooxml` members of the same name.
 */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class WorkbookOoxml {
  /** @internal */
  constructor() {
    throw new TypeError('read Workbook.ooxml instead');
  }
}

/**
 * The OOXML-level layer of a sheet, `Sheet.ooxml`: the cell, row and column XF indices, the range-XF
 * remap and the sheet-level XML fragments. Members are the engine's `Sheet.ooxml` members of the
 * same name.
 */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class SheetOoxml {
  /** @internal */
  constructor() {
    throw new TypeError('read Sheet.ooxml instead');
  }
}

export const WORKBOOK_OOXML_MEMBERS = {
  read: [
    'getCellXf',
    'listCellXfs',
    'getCellStyleXf',
    'listCellStyleXfs',
    'getFont',
    'listFonts',
    'getFill',
    'listFills',
    'getBorder',
    'listBorders',
    'getNumberFormat',
    'listNumberFormats',
    'getDxf',
    'listDxfs',
    'hasPassthrough',
    'listPassthroughs',
  ],
  change: {
    format: [
      'addCellXf',
      'setCellXf',
      'addCellStyleXf',
      'setCellStyleXf',
      'addFont',
      'setFont',
      'addFill',
      'setFill',
      'addBorder',
      'setBorder',
      'addNumberFormat',
      'addDxf',
      'setDxf',
    ],
  },
} as const satisfies MemberSpec<WorkbookContext>;

export const SHEET_OOXML_MEMBERS = {
  read: [
    'getCellXfIndex',
    'getRowXfIndex',
    'getColumnXfIndex',
    'getEffectiveStyle',
    'getRangeXfBases',
    'getRangeXfBasesExcluding',
    'snapshotRangeXfs',
    'getAutoFilterXml',
    'getHeaderFooterXml',
    'getPageMarginsXml',
    'getPageSetupXml',
    'getPrintOptionsXml',
    'getSheetPrXml',
  ],
  change: {
    format: [
      'setCellXfIndex',
      'resetCellXfIndex',
      'setRowXfIndex',
      'resetRowXfIndex',
      'setColumnXfIndex',
      'resetColumnXfIndex',
      'setRangeXfIndex',
      'remapRangeXfs',
      'restoreRangeXfs',
    ],
    print: ['setHeaderFooterXml', 'setPageMarginsXml', 'setPageSetupXml', 'setPrintOptionsXml'],
    autoFilter: ['setAutoFilterXml'],
    view: ['setSheetPrXml'],
  },
} as const satisfies MemberSpec<SheetContext>;

type WorkbookOoxmlChanges = typeof WORKBOOK_OOXML_MEMBERS.change;
type WorkbookOoxmlMemberName =
  | (typeof WORKBOOK_OOXML_MEMBERS.read)[number]
  | WorkbookOoxmlChanges[keyof WorkbookOoxmlChanges][number];
type SheetOoxmlChanges = typeof SHEET_OOXML_MEMBERS.change;
type SheetOoxmlMemberName =
  | (typeof SHEET_OOXML_MEMBERS.read)[number]
  | SheetOoxmlChanges[keyof SheetOoxmlChanges][number];

export interface WorkbookOoxml extends Pick<EngineWorkbookOoxml, WorkbookOoxmlMemberName> {}
export interface SheetOoxml extends Pick<EngineSheetOoxml, SheetOoxmlMemberName> {}
