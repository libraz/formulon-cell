import {
  type FunctionCatalogEntry,
  type FunctionCatalogMetadata,
  functionSyntax,
} from '../commands/function-categories.js';
import { FUNCTION_SIGNATURES } from '../commands/refs.js';
import { functionDescription } from './function-catalog-text.js';

export interface FormulaHelperDetails {
  signature?: string;
  description?: string;
}

export interface FormulaHelperCatalogOptions {
  localeTag: string;
  listSeparator?: string;
  getMetadata: (canonicalName: string, localeTag: string) => FunctionCatalogMetadata | null;
  localizeName: (canonicalName: string) => string;
  getCustomFunctionArgs?: (canonicalName: string) => readonly string[] | undefined;
}

const arityFromLabels = (
  labels: readonly string[] | undefined,
): { minArity: number; maxArity: number | null; argumentLabels: readonly string[] } => {
  if (!labels) return { minArity: 0, maxArity: null, argumentLabels: [] };
  const argumentLabels = labels.filter((label) => label !== '...');
  return {
    minArity: argumentLabels.filter((label) => !label.startsWith('[')).length,
    maxArity: labels.includes('...') ? null : argumentLabels.length,
    argumentLabels,
  };
};

/** Resolve help for one active candidate only; callers provide the live
 * workbook locale and metadata reader so the full catalog is never scanned. */
export function resolveFormulaHelperDetails(
  canonicalName: string,
  options: FormulaHelperCatalogOptions,
): FormulaHelperDetails | undefined {
  const canonical = canonicalName.trim().toLocaleUpperCase();
  if (!canonical) return undefined;

  const metadata = options.getMetadata(canonical, options.localeTag);
  const staticLabels = FUNCTION_SIGNATURES[canonical];
  const customLabels = staticLabels ? undefined : options.getCustomFunctionArgs?.(canonical);
  const fallbackLabels = staticLabels ?? customLabels;
  const arity = arityFromLabels(fallbackLabels);
  const localized = options.localizeName(canonical) || canonical;

  const entry: FunctionCatalogEntry = {
    canonicalName: canonical,
    displayName: metadata?.localizedName ?? localized,
    minArity: metadata?.minArity ?? arity.minArity,
    maxArity: metadata === null ? arity.maxArity : metadata.maxArity,
    argumentLabels: arity.argumentLabels,
    ...(metadata?.signatureTemplate === undefined
      ? {}
      : { signatureTemplate: metadata.signatureTemplate }),
    ...(metadata?.description === undefined ? {} : { description: metadata.description }),
    ...(metadata?.availability === undefined ? {} : { availability: metadata.availability }),
  };

  const hasSignature = metadata?.signatureTemplate !== undefined || fallbackLabels !== undefined;
  const signature = hasSignature ? functionSyntax(entry, options.listSeparator) : undefined;
  const description = functionDescription(entry, options.localeTag);
  if (!signature && !description) return undefined;
  return {
    ...(signature ? { signature } : {}),
    ...(description ? { description } : {}),
  };
}
