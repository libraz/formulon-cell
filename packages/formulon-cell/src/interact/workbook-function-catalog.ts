import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import type { CatalogLocaleOrdinal } from '../commands/function-categories.js';
import { FUNCTION_SIGNATURES, findActiveSignature } from '../commands/refs.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { ArgHelperDeps } from './arg-helper.js';
import { type AutocompleteDeps, completeFunctionNames } from './autocomplete.js';
import { collectLambdaFunctions } from './autocomplete-names.js';
import { caretSplitsToken } from './autocomplete-suggestions.js';
import { resolveFormulaHelperDetails } from './formula-helper-catalog.js';
import {
  catalogLocaleOrdinal,
  catalogLocaleTag,
  functionListSeparator,
  nativeFunctionName,
} from './function-catalog-text.js';

type CustomFunctionArgs = (name: string) => readonly string[] | undefined;

/** Function names are localized in the workbook's input locale (the language
 *  formulas are typed in), not the UI locale. */
const inputCatalogLocale = (wb: WorkbookHandle): { tag: string; ordinal: CatalogLocaleOrdinal } => {
  const tag = catalogLocaleTag(inputLocaleForWorkbook(wb) ?? 'en-US');
  return { tag, ordinal: catalogLocaleOrdinal(tag) };
};

const resolveFunctionName = (wb: WorkbookHandle, rawName: string): string =>
  wb.canonicalizeFunctionName(rawName, inputCatalogLocale(wb).ordinal) ||
  rawName.toLocaleUpperCase();

/** The custom functions a formula can call from the edited sheet: the host's registered
 *  hints, then the workbook's LAMBDA names with their parameters as argument labels. */
export function workbookCustomFunctions(
  getWb: () => WorkbookHandle,
  getSheet: () => number,
  hints: { list: () => readonly string[]; args: CustomFunctionArgs },
): Required<Pick<AutocompleteDeps, 'getCustomFunctions' | 'getCustomFunctionArgs'>> {
  const lambdas = () => collectLambdaFunctions(getWb().definedNames(), getSheet());
  return {
    getCustomFunctions: () => [
      ...hints.list(),
      ...Array.from(lambdas().values(), (lambda) => lambda.name),
    ],
    getCustomFunctionArgs: (name) =>
      hints.args(name) ?? lambdas().get(name.toLocaleLowerCase())?.params,
  };
}

/** Autocomplete deps backed by the workbook's function registry, shared by
 *  the in-cell editor and the formula bar. `getWb` is read on every call so a
 *  swapped workbook is picked up without re-attaching. */
export function workbookAutocompleteCatalog(
  getWb: () => WorkbookHandle,
  getCustomFunctionArgs: CustomFunctionArgs,
): Required<
  Pick<
    AutocompleteDeps,
    | 'getFunctionNames'
    | 'getFunctionDisplayNames'
    | 'resolveFunctionName'
    | 'getListSeparator'
    | 'getFunctionDetails'
  >
> {
  return {
    getFunctionNames: () => getWb().functionNames(),
    getFunctionDisplayNames: () => {
      const wb = getWb();
      const names = wb.functionNames();
      if (!names) return null;
      const { ordinal } = inputCatalogLocale(wb);
      return names.map((name) => nativeFunctionName(wb, name, ordinal));
    },
    resolveFunctionName: (rawName) => resolveFunctionName(getWb(), rawName),
    getListSeparator: () => functionListSeparator(getWb()),
    getFunctionDetails: (canonicalName) => {
      const wb = getWb();
      const { tag, ordinal } = inputCatalogLocale(wb);
      return resolveFormulaHelperDetails(canonicalName, {
        localeTag: tag,
        listSeparator: functionListSeparator(wb),
        getMetadata: (name, locale) => wb.functionMetadataForLocale(name, locale),
        localizeName: (name) => nativeFunctionName(wb, name, ordinal),
        getCustomFunctionArgs,
      });
    },
  };
}

/** Argument-helper deps backed by the workbook's function registry. */
export function workbookArgHelperCatalog(
  getWb: () => WorkbookHandle,
): Required<
  Pick<ArgHelperDeps, 'getListSeparator' | 'getFunctionDisplayName' | 'resolveFunctionName'>
> {
  return {
    getListSeparator: () => functionListSeparator(getWb()),
    getFunctionDisplayName: (canonicalName) => {
      const wb = getWb();
      const { tag, ordinal } = inputCatalogLocale(wb);
      return (
        wb.functionMetadataForLocale(canonicalName, tag)?.localizedName ||
        nativeFunctionName(wb, canonicalName, ordinal)
      );
    },
    resolveFunctionName: (rawName) => resolveFunctionName(getWb(), rawName),
  };
}

const NO_HINTS = { list: () => [], args: () => undefined };

/** A function call the caret sits in: its canonical name, argument labels and the argument the
 *  caret is in. */
export interface FunctionCallAtCaret {
  readonly name: string;
  readonly args: readonly string[];
  readonly argIndex: number;
}

/** What a function call is, for a tooltip: its canonical and localized names, argument labels
 *  and description. */
export interface FunctionDescription {
  readonly name: string;
  readonly localizedName: string;
  readonly args: readonly string[];
  readonly description: string;
}

/**
 * The formula editor's function completion and argument help for `wb`, as the in-cell editor
 * and the formula bar compute them, without a mount's formula hints. `sheet` is the sheet the
 * formula is on, whose sheet-scoped LAMBDA names it may call; -1 limits them to the workbook's.
 */
export function workbookFunctionHelp(wb: WorkbookHandle, sheet: number) {
  const custom = workbookCustomFunctions(
    () => wb,
    () => sheet,
    NO_HINTS,
  );
  const catalog = workbookAutocompleteCatalog(() => wb, custom.getCustomFunctionArgs);
  const argCatalog = workbookArgHelperCatalog(() => wb);
  const isBuiltIn = (name: string): boolean => (wb.functionNames() ?? []).includes(name);
  return {
    suggest(text: string, caret: number) {
      if (caretSplitsToken(text, caret)) return null;
      return completeFunctionNames(text, caret, { ...catalog, ...custom });
    },
    callAt(text: string, caret: number): FunctionCallAtCaret | null {
      const sig = findActiveSignature(text, caret, {
        listSeparator: argCatalog.getListSeparator(),
        resolveKnownName: argCatalog.resolveFunctionName,
        getCustomFunctionArgs: custom.getCustomFunctionArgs,
        isKnownFunction: isBuiltIn,
      });
      if (!sig) return null;
      return { name: sig.name, args: sig.args, argIndex: sig.activeArgIndex };
    },
    describe(rawName: string): FunctionDescription | null {
      const name = argCatalog.resolveFunctionName(rawName.trim());
      if (!name) return null;
      const builtIn = isBuiltIn(name);
      const lambdaArgs = builtIn ? undefined : custom.getCustomFunctionArgs(name);
      if (!builtIn && !lambdaArgs) return null;
      return {
        name,
        localizedName: builtIn ? (argCatalog.getFunctionDisplayName(name) ?? name) : name,
        args: FUNCTION_SIGNATURES[name] ?? lambdaArgs ?? [],
        description: builtIn ? (catalog.getFunctionDetails(name)?.description ?? '') : '',
      };
    },
  };
}
