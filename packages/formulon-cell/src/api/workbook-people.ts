import type { Workbook as EngineWorkbook } from '@libraz/formulon';
import type { MemberSpec, WorkbookContext } from './facade-context.js';

/**
 * The workbook's person list (the authors threaded comments and mentions name) and its
 * external-link records, as engine `Workbook` members of the same name.
 */
export const WORKBOOK_PEOPLE_MEMBERS = {
  read: ['listPeople', 'getPerson', 'listExternalLinks'],
  change: { threadedComments: ['addPerson', 'removePerson'] },
} as const satisfies MemberSpec<WorkbookContext>;

type PeopleMemberName =
  | (typeof WORKBOOK_PEOPLE_MEMBERS.read)[number]
  | (typeof WORKBOOK_PEOPLE_MEMBERS.change.threadedComments)[number];

declare module './workbook.js' {
  interface Workbook extends Pick<EngineWorkbook, PeopleMemberName> {}
}
