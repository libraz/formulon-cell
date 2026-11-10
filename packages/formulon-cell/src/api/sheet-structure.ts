import type { Sheet as EngineSheet, Workbook as EngineWorkbook } from '@libraz/formulon';
import type { MemberSpec, SheetContext, WorkbookContext } from './facade-context.js';

/**
 * Rows, columns, merges, sizes, view, hyperlinks, comments, AutoFilter and sparklines: engine
 * `Sheet` members of the same name. Each write announces the part of the sheet it changed.
 */
export const SHEET_STRUCTURE_MEMBERS = {
  read: [
    'listMerges',
    'getMergesInRange',
    'listRowOverrides',
    'listColumnOverrides',
    'getRowHeightPt',
    'getColumnWidthPt',
    'getCellRectPt',
    'getWidthModel',
    'convertColumnCharsToPt',
    'convertColumnPtToChars',
    'recalcViewport',
    'getView',
    'getSelection',
    'getTabColor',
    'getVisibility',
    'getHyperlink',
    'listHyperlinks',
    'getComment',
    'listComments',
    'getCommentVisible',
    'getThreadedComment',
    'listThreadedComments',
    'getAutoFilter',
    'evaluateAutoFilter',
    'getSparklineGroup',
    'listSparklineGroups',
  ],
  change: {
    layout: [
      'insertRows',
      'deleteRows',
      'insertColumns',
      'deleteColumns',
      'insertCells',
      'deleteCells',
      'addMerge',
      'removeMerge',
      'clearMerges',
      'setRowHeight',
      'setRowAutoHeight',
      'resetRowHeight',
      'setRowHidden',
      'setRowOutline',
      'setRowCollapsed',
      'setColumnWidth',
      'resetColumnWidth',
      'setColumnHidden',
      'setColumnOutline',
      'setColumnCollapsed',
    ],
    view: ['updateView', 'setSelection', 'setTabColor', 'setVisibility'],
    format: ['setHyperlink', 'clearHyperlinks'],
    comments: ['setComment', 'setCommentVisible'],
    threadedComments: ['addThreadedComment', 'updateThreadedComment', 'removeThreadedComment'],
    sparklineGroups: ['addSparklineGroup', 'updateSparklineGroup', 'removeSparklineGroup'],
    autoFilter: ['setAutoFilter', 'applyAutoFilter', 'clearAutoFilterCriteria'],
  },
} as const satisfies MemberSpec<SheetContext>;

/** Defined names: engine `Workbook` members of the same name. */
export const WORKBOOK_NAME_MEMBERS = {
  read: ['getDefinedName', 'listDefinedNames'],
  change: { names: ['setDefinedName', 'removeDefinedName'] },
} as const satisfies MemberSpec<WorkbookContext>;

type Changes = typeof SHEET_STRUCTURE_MEMBERS.change;
type StructureMemberName =
  | (typeof SHEET_STRUCTURE_MEMBERS.read)[number]
  | Changes[keyof Changes][number];
type NameMemberName =
  | (typeof WORKBOOK_NAME_MEMBERS.read)[number]
  | (typeof WORKBOOK_NAME_MEMBERS.change.names)[number];

declare module './workbook.js' {
  interface Sheet extends Pick<EngineSheet, StructureMemberName> {}
  interface Workbook extends Pick<EngineWorkbook, NameMemberName> {}
}
