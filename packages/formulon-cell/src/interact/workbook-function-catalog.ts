import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import type { CatalogLocaleOrdinal } from '../commands/function-categories.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { ArgHelperDeps } from './arg-helper.js';
import type { AutocompleteDeps } from './autocomplete.js';
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

/** Autocomplete deps backed by the workbook's function registry, shared by
 *  the in-cell editor and the formula bar. `getWb` is read on every call so a
 *  swapped workbook is picked up without re-attaching. */
export function workbookAutocompleteCatalog(
  getWb: () => WorkbookHandle,
  getCustomFunctionArgs: CustomFunctionArgs,
): Pick<
  AutocompleteDeps,
  | 'getFunctionNames'
  | 'getFunctionDisplayNames'
  | 'resolveFunctionName'
  | 'getListSeparator'
  | 'getFunctionDetails'
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
): Pick<ArgHelperDeps, 'getListSeparator' | 'getFunctionDisplayName' | 'resolveFunctionName'> {
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
