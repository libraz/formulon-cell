import type { CalcMode, Sheet as EngineSheet, Workbook as EngineWorkbook } from '@libraz/formulon';
import {
  type FunctionCallAtCaret,
  type FunctionDescription,
  workbookFunctionHelp,
} from '../interact/workbook-function-catalog.js';
import type { MemberSpec, WorkbookContext } from './facade-context.js';
import type { SheetRef } from './workbook.js';

/** Where a formula is typed: the sheet whose sheet-scoped LAMBDA names it may call. */
export interface FunctionHelpOptions {
  /** The sheet the formula is on. Omitted, only the workbook's own LAMBDA names are offered. */
  readonly sheet?: SheetRef;
}

/** Options of `Workbook.suggestFunctions`. */
export interface SuggestFunctionsOptions extends FunctionHelpOptions {
  /** The most names to return. Default 8. */
  readonly max?: number;
}

/** The functions `Workbook.suggestFunctions` offers for the name being typed. */
export interface FunctionSuggestions {
  /** The partial name before the caret. */
  readonly token: string;
  /** Where `token` starts in the formula, as a UTF-16 offset. */
  readonly tokenStart: number;
  /** Matching names as they are typed in the workbook's language, best first. */
  readonly names: readonly string[];
}

export type { FunctionCallAtCaret, FunctionDescription };

const DEFAULT_SUGGESTIONS = 8;

const helpFor = (ctx: WorkbookContext, options: FunctionHelpOptions | undefined) =>
  workbookFunctionHelp(
    ctx.handle,
    options?.sheet === undefined
      ? -1
      : ctx.engine.sheet(ctx.engineOf(options.sheet) as EngineSheet | number | string).index,
  );

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
    suggestFunctions: (
      ctx: WorkbookContext,
      formula: string,
      caret: number,
      options?: SuggestFunctionsOptions,
    ): FunctionSuggestions | null => {
      const found = helpFor(ctx, options).suggest(formula, caret);
      if (!found) return null;
      const names = found.names.slice(0, options?.max ?? DEFAULT_SUGGESTIONS);
      return names.length > 0 ? { token: found.token, tokenStart: found.tokenStart, names } : null;
    },
    functionCallAt: (
      ctx: WorkbookContext,
      formula: string,
      caret: number,
      options?: FunctionHelpOptions,
    ): FunctionCallAtCaret | null => helpFor(ctx, options).callAt(formula, caret),
    describeFunction: (
      ctx: WorkbookContext,
      name: string,
      options?: FunctionHelpOptions,
    ): FunctionDescription | null => helpFor(ctx, options).describe(name),
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
    /**
     * The functions to offer while `formula` is typed with the caret at `caret` (a UTF-16
     * offset), as the formula editor's autocomplete lists them: built-in functions in the
     * workbook's language and the workbook's LAMBDA names. A mounted spreadsheet's
     * `formulaHints` add to its own editor only. Null when the caret is not at a function name.
     * @example
     * workbook.suggestFunctions('=SU', 3)?.names; // ['SUBSTITUTE', 'SUBTOTAL', 'SUM', ...]
     */
    suggestFunctions(
      formula: string,
      caret: number,
      options?: SuggestFunctionsOptions,
    ): FunctionSuggestions | null;
    /**
     * The function call the caret (a UTF-16 offset) sits in, as the formula editor's argument
     * tooltip shows it: the innermost call with known arguments. Null outside any call, or in a
     * function with no argument labels.
     * @example
     * workbook.functionCallAt('=SUM(A1,', 8); // { name: 'SUM', args: ['number1', ...], argIndex: 1 }
     */
    functionCallAt(
      formula: string,
      caret: number,
      options?: FunctionHelpOptions,
    ): FunctionCallAtCaret | null;
    /**
     * A built-in function or LAMBDA name, given canonical or in the workbook's language: its
     * names, argument labels and description, in the workbook's locale, as the in-cell editor
     * shows them; there is no locale option. The description comes from the workbook's
     * `functionMetadata`, else the engine's metadata, else the built-in text; a LAMBDA has its
     * parameters as arguments and no description. Null for any other name.
     */
    describeFunction(name: string, options?: FunctionHelpOptions): FunctionDescription | null;
  }
}
