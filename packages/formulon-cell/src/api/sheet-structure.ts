import type { Sheet as EngineSheet } from '@libraz/formulon';
import type { MemberSpec, SheetContext } from './facade-context.js';

/** Rows, columns, merges and the other structure of a sheet: engine `Sheet` members of the same name. */
export const SHEET_STRUCTURE_MEMBERS = {
  read: ['listMerges'],
  change: { layout: ['addMerge', 'removeMerge'] },
} as const satisfies MemberSpec<SheetContext>;

type StructureMemberName =
  | (typeof SHEET_STRUCTURE_MEMBERS.read)[number]
  | (typeof SHEET_STRUCTURE_MEMBERS.change.layout)[number];

declare module './workbook.js' {
  interface Sheet extends Pick<EngineSheet, StructureMemberName> {}
}
