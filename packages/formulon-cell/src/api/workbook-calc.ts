import type { CalcMode, Workbook as EngineWorkbook } from '@libraz/formulon';
import type { MemberSpec, WorkbookContext } from './facade-context.js';

/**
 * Calculation settings of the workbook (`<calcPr>`), as engine `Workbook` members of the same
 * name, and the catalog of built-in functions. The calc mode goes through the workbook's own
 * policy: leaving Manual recalculates what the edits made under Manual left dirty.
 */
export const WORKBOOK_CALC_MEMBERS = {
  read: ['getCalcMode', 'getIterative', 'getProfileId'],
  change: { calc: ['updateIterative'] },
  methods: {
    setCalcMode: (ctx: WorkbookContext, mode: CalcMode) =>
      ctx.change('calc', null, () => ctx.handle.changeCalcMode(mode)),
    listFunctionNames: (ctx: WorkbookContext) => ctx.handle.engineModule().listFunctionNames(),
  },
} as const satisfies MemberSpec<WorkbookContext>;

type CalcMemberName =
  | (typeof WORKBOOK_CALC_MEMBERS.read)[number]
  | (typeof WORKBOOK_CALC_MEMBERS.change.calc)[number];

declare module './workbook.js' {
  interface Workbook extends Pick<EngineWorkbook, CalcMemberName> {
    /** Sets the calc mode. Leaving `CalcMode.Manual` recalculates the cells left dirty while it
     *  was on. While mounted, the ribbon shows the new mode. */
    setCalcMode(mode: CalcMode): void;
    /** The canonical names of every built-in function, in ascending order. */
    listFunctionNames(): readonly string[];
  }
}
