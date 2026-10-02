import { FUNCTION_SIGNATURES } from './refs.js';

/** Function families exposed by the function picker and the Mac ribbon.
 *
 * `all` and `recent` are picker views rather than catalog families. The
 * remaining names mirror the families shown by Excel's Function Library;
 * empty families are retained so hosts can add signatures without changing
 * the public category type.
 */
export type FunctionCategory =
  | 'all'
  | 'recent'
  | 'logical'
  | 'lookup'
  | 'text'
  | 'datetime'
  | 'math'
  | 'financial'
  | 'dynamicArray'
  | 'statistical'
  | 'engineering'
  | 'information'
  | 'database'
  | 'compatibility'
  | 'cube'
  | 'web';

export type CatalogFunctionCategory = Exclude<FunctionCategory, 'all' | 'recent'>;

export interface FunctionCatalogReader {
  functionNames?: () => readonly string[] | null;
  functionMetadata?: (name: string, locale: 0 | 1) => FunctionCatalogMetadata | null;
}

export interface FunctionCatalogMetadata {
  name?: string;
  minArity: number;
  maxArity: number | null;
  localizedName?: string;
  signatureTemplate?: string;
  description?: string;
}

export interface FunctionCatalogEntry {
  canonicalName: string;
  displayName: string;
  minArity: number;
  maxArity: number | null;
  argumentLabels: readonly string[];
  signatureTemplate?: string;
  description?: string;
}

export interface FunctionCatalogSnapshot {
  names: readonly string[];
  knownNames: ReadonlySet<string>;
  entries: ReadonlyMap<string, FunctionCatalogEntry>;
  source: 'engine' | 'fallback';
}

/** Stable family membership. A function may occur in more than one family,
 * just as it does in Excel's picker (for example AVERAGE is both Math and
 * Statistical). Keep the catalog names upper-cased to match
 * `FUNCTION_SIGNATURES`. */
export const FUNCTION_CATEGORY_NAMES: Readonly<Record<CatalogFunctionCategory, readonly string[]>> =
  {
    logical: ['IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'TRUE', 'FALSE'],
    lookup: [
      'VLOOKUP',
      'HLOOKUP',
      'XLOOKUP',
      'INDEX',
      'MATCH',
      'XMATCH',
      'OFFSET',
      'INDIRECT',
      'CHOOSE',
      'ROW',
      'COLUMN',
      'ROWS',
      'COLUMNS',
    ],
    text: [
      'CONCATENATE',
      'CONCAT',
      'TEXTJOIN',
      'TEXTSPLIT',
      'TEXTBEFORE',
      'TEXTAFTER',
      'LEFT',
      'RIGHT',
      'MID',
      'LEN',
      'UPPER',
      'LOWER',
      'PROPER',
      'TRIM',
      'SUBSTITUTE',
      'REPLACE',
      'FIND',
      'SEARCH',
      'TEXT',
      'VALUE',
      'NUMBERVALUE',
    ],
    datetime: [
      'TODAY',
      'NOW',
      'DATE',
      'YEAR',
      'MONTH',
      'DAY',
      'HOUR',
      'MINUTE',
      'SECOND',
      'WEEKDAY',
      'EOMONTH',
      'DATEDIF',
      'NETWORKDAYS',
      'WORKDAY',
    ],
    math: [
      'SUM',
      'AVERAGE',
      'COUNT',
      'COUNTA',
      'COUNTIF',
      'COUNTIFS',
      'SUMIF',
      'SUMIFS',
      'AVERAGEIF',
      'AVERAGEIFS',
      'MIN',
      'MAX',
      'MEDIAN',
      'ROUND',
      'ROUNDUP',
      'ROUNDDOWN',
      'CEILING',
      'FLOOR',
      'INT',
      'MOD',
      'ABS',
      'POWER',
      'SQRT',
      'EXP',
      'LN',
      'LOG',
      'LOG10',
    ],
    financial: ['PMT', 'PV', 'FV', 'NPV', 'IRR', 'RATE', 'NPER'],
    dynamicArray: [
      'TRANSPOSE',
      'UNIQUE',
      'SORT',
      'SORTBY',
      'FILTER',
      'SEQUENCE',
      'RANDARRAY',
      'VSTACK',
      'HSTACK',
      'TOROW',
      'TOCOL',
      'WRAPROWS',
      'WRAPCOLS',
      'CHOOSEROWS',
      'CHOOSECOLS',
      'TAKE',
      'DROP',
      'EXPAND',
      'LAMBDA',
      'LET',
      'MAP',
      'REDUCE',
      'SCAN',
      'BYROW',
      'BYCOL',
      'MAKEARRAY',
      'GROUPBY',
      'PIVOTBY',
      'PERCENTOF',
      'IMAGE',
    ],
    statistical: [
      'AVERAGE',
      'COUNT',
      'COUNTA',
      'COUNTIF',
      'COUNTIFS',
      'AVERAGEIF',
      'AVERAGEIFS',
      'MAX',
      'MEDIAN',
      'MIN',
      'SUM',
      'SUMIF',
      'SUMIFS',
    ],
    engineering: [],
    information: [],
    database: [],
    compatibility: [],
    cube: [],
    web: [],
  };

/** Return the functions in one picker family that this engine actually
 * knows. The result is sorted and never exposes an empty/unknown ribbon
 * leaf. */
export const supportedFunctionNames = (
  category: CatalogFunctionCategory,
  liveNames?: ReadonlySet<string>,
): string[] =>
  FUNCTION_CATEGORY_NAMES[category]
    .filter((name) => (liveNames === undefined ? name in FUNCTION_SIGNATURES : liveNames.has(name)))
    .slice()
    .sort();

/** Return every known function in the engine's catalog in picker order. */
export const allFunctionNames = (liveNames?: ReadonlySet<string>): string[] =>
  liveNames === undefined ? Object.keys(FUNCTION_SIGNATURES).sort() : [...liveNames].sort();

const staticArity = (
  signature: readonly string[] | undefined,
): { minArity: number; maxArity: number | null; argumentLabels: readonly string[] } => {
  if (!signature) return { minArity: 0, maxArity: null, argumentLabels: [] };
  const argumentLabels = signature.filter((label) => label !== '...');
  const minArity = argumentLabels.filter((label) => !label.startsWith('[')).length;
  return {
    minArity,
    maxArity: signature.includes('...') ? null : argumentLabels.length,
    argumentLabels,
  };
};

/** Build one immutable view of the workbook's recognized functions. A valid
 * empty engine result remains authoritative; static signatures are used only
 * when the reader is absent or reports that the capability is unavailable. */
export const buildFunctionCatalog = (
  reader: FunctionCatalogReader | null,
  locale: 0 | 1,
): FunctionCatalogSnapshot => {
  const live = reader?.functionNames?.() ?? null;
  const source = live === null ? 'fallback' : 'engine';
  const names = (live === null ? Object.keys(FUNCTION_SIGNATURES) : live)
    .map((name) => name.trim().toUpperCase())
    .filter((name) => name.length > 0);
  const uniqueNames = [...new Set(names)].sort();
  const entries = new Map<string, FunctionCatalogEntry>();
  for (const canonicalName of uniqueNames) {
    const fallback = staticArity(FUNCTION_SIGNATURES[canonicalName]);
    const metadata = reader?.functionMetadata?.(canonicalName, locale) ?? null;
    const minArity = metadata === null ? fallback.minArity : metadata.minArity;
    const maxArity = metadata === null ? fallback.maxArity : metadata.maxArity;
    const displayName = metadata?.localizedName ?? canonicalName;
    entries.set(canonicalName, {
      canonicalName,
      displayName,
      minArity: Math.max(0, minArity),
      maxArity: maxArity === null ? null : Math.max(0, maxArity),
      argumentLabels: fallback.argumentLabels,
      ...(metadata?.signatureTemplate === undefined
        ? {}
        : { signatureTemplate: metadata.signatureTemplate }),
      ...(metadata?.description === undefined ? {} : { description: metadata.description }),
    });
  }
  return {
    names: Object.freeze(uniqueNames),
    knownNames: new Set(uniqueNames),
    entries,
    source,
  };
};
