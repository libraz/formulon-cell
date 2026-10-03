import { FUNCTION_SIGNATURES } from './refs.js';

/** Function families exposed by the function picker and the Mac ribbon.
 *
 * `all` and `recent` are picker views rather than catalog families. The
 * remaining names follow the standard spreadsheet function-library families;
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
  /** Raw engine availability class. Only exactly 3 is unavailable. */
  availability?: number;
}

export interface FunctionCatalogEntry {
  canonicalName: string;
  displayName: string;
  minArity: number;
  maxArity: number | null;
  argumentLabels: readonly string[];
  signatureTemplate?: string;
  description?: string;
  availability?: number;
}

/** Display syntax for a catalog entry: the metadata template when present,
 * otherwise `NAME(labels)` with a trailing `...` for unbounded variadics —
 * `NAME(a, [b], ...)` when labelled, `NAME(...)` when not. */
export const functionSyntax = (entry: FunctionCatalogEntry): string => {
  if (entry.signatureTemplate) return entry.signatureTemplate;
  const parts = [...entry.argumentLabels];
  if (entry.maxArity === null) parts.push('...');
  return `${entry.displayName}(${parts.join(', ')})`;
};

/** Whether an engine function is a visible-but-non-insertable stub. */
export const isFunctionUnavailableForInsertion = (
  availability: number | null | undefined,
): boolean => availability === 3;

export interface FunctionCatalogSnapshot {
  names: readonly string[];
  knownNames: ReadonlySet<string>;
  entries: ReadonlyMap<string, FunctionCatalogEntry>;
  source: 'engine' | 'fallback';
}

/** Stable family membership for the native engine's function catalog. Each
 * function belongs to exactly one family; `dynamicArray` is an additional
 * convenience view and may overlap one. `all` and `recent` are picker views
 * rather than catalog families. Legacy names (e.g. NORMINV) live under
 * `compatibility`, their current replacements (NORM.INV) under `statistical`. */
export const FUNCTION_CATEGORY_NAMES: Readonly<Record<CatalogFunctionCategory, readonly string[]>> =
  {
    logical: [
      'AND',
      'BYCOL',
      'BYROW',
      'FALSE',
      'IF',
      'IFERROR',
      'IFNA',
      'IFS',
      'LAMBDA',
      'LET',
      'MAKEARRAY',
      'MAP',
      'NOT',
      'OR',
      'REDUCE',
      'SCAN',
      'SWITCH',
      'TRUE',
      'XOR',
    ],
    lookup: [
      'ADDRESS',
      'AREAS',
      'CHOOSE',
      'CHOOSECOLS',
      'CHOOSEROWS',
      'COLUMN',
      'COLUMNS',
      'DROP',
      'EXPAND',
      'FILTER',
      'FORMULATEXT',
      'GETPIVOTDATA',
      'GROUPBY',
      'HLOOKUP',
      'HSTACK',
      'HYPERLINK',
      'IMAGE',
      'INDEX',
      'INDIRECT',
      'LOOKUP',
      'MATCH',
      'OFFSET',
      'PIVOTBY',
      'ROW',
      'ROWS',
      'RTD',
      'SORT',
      'SORTBY',
      'TAKE',
      'TOCOL',
      'TOROW',
      'TRANSPOSE',
      'TRIMRANGE',
      'UNIQUE',
      'VLOOKUP',
      'VSTACK',
      'WRAPCOLS',
      'WRAPROWS',
      'XLOOKUP',
      'XMATCH',
    ],
    text: [
      'ARRAYTOTEXT',
      'ASC',
      'BAHTTEXT',
      'CHAR',
      'CLEAN',
      'CODE',
      'CONCAT',
      'CONCATENATE',
      'DBCS',
      'DETECTLANGUAGE',
      'DOLLAR',
      'EXACT',
      'FIND',
      'FINDB',
      'FIXED',
      'JIS',
      'LEFT',
      'LEFTB',
      'LEN',
      'LENB',
      'LOWER',
      'MID',
      'MIDB',
      'NUMBERVALUE',
      'PHONETIC',
      'PROPER',
      'REGEXEXTRACT',
      'REGEXREPLACE',
      'REGEXTEST',
      'REPLACE',
      'REPLACEB',
      'REPT',
      'RIGHT',
      'RIGHTB',
      'SEARCH',
      'SEARCHB',
      'SUBSTITUTE',
      'T',
      'TEXT',
      'TEXTAFTER',
      'TEXTBEFORE',
      'TEXTJOIN',
      'TEXTSPLIT',
      'TRANSLATE',
      'TRIM',
      'UNICHAR',
      'UNICODE',
      'UPPER',
      'VALUE',
      'VALUETOTEXT',
    ],
    datetime: [
      'DATE',
      'DATEDIF',
      'DATEVALUE',
      'DAY',
      'DAYS',
      'DAYS360',
      'EDATE',
      'EOMONTH',
      'HOUR',
      'ISOWEEKNUM',
      'MINUTE',
      'MONTH',
      'NETWORKDAYS',
      'NETWORKDAYS.INTL',
      'NOW',
      'SECOND',
      'TIME',
      'TIMEVALUE',
      'TODAY',
      'WEEKDAY',
      'WEEKNUM',
      'WORKDAY',
      'WORKDAY.INTL',
      'YEAR',
      'YEARFRAC',
    ],
    math: [
      'ABS',
      'ACOS',
      'ACOSH',
      'ACOT',
      'ACOTH',
      'AGGREGATE',
      'ARABIC',
      'ASIN',
      'ASINH',
      'ATAN',
      'ATAN2',
      'ATANH',
      'BASE',
      'CEILING.MATH',
      'CEILING.PRECISE',
      'COMBIN',
      'COMBINA',
      'COS',
      'COSH',
      'COT',
      'COTH',
      'CSC',
      'CSCH',
      'DECIMAL',
      'DEGREES',
      'EVEN',
      'EXP',
      'FACT',
      'FACTDOUBLE',
      'FLOOR.MATH',
      'FLOOR.PRECISE',
      'GCD',
      'INT',
      'ISO.CEILING',
      'LCM',
      'LN',
      'LOG',
      'LOG10',
      'MDETERM',
      'MINVERSE',
      'MMULT',
      'MOD',
      'MROUND',
      'MULTINOMIAL',
      'MUNIT',
      'ODD',
      'PERCENTOF',
      'PI',
      'POWER',
      'PRODUCT',
      'QUOTIENT',
      'RADIANS',
      'RAND',
      'RANDARRAY',
      'RANDBETWEEN',
      'ROMAN',
      'ROUND',
      'ROUNDDOWN',
      'ROUNDUP',
      'SEC',
      'SECH',
      'SEQUENCE',
      'SERIESSUM',
      'SIGN',
      'SIN',
      'SINH',
      'SQRT',
      'SQRTPI',
      'SUBTOTAL',
      'SUM',
      'SUMIF',
      'SUMIFS',
      'SUMPRODUCT',
      'SUMSQ',
      'SUMX2MY2',
      'SUMX2PY2',
      'SUMXMY2',
      'TAN',
      'TANH',
      'TRUNC',
    ],
    financial: [
      'ACCRINT',
      'ACCRINTM',
      'AMORDEGRC',
      'AMORLINC',
      'COUPDAYBS',
      'COUPDAYS',
      'COUPDAYSNC',
      'COUPNCD',
      'COUPNUM',
      'COUPPCD',
      'CUMIPMT',
      'CUMPRINC',
      'DB',
      'DDB',
      'DISC',
      'DOLLARDE',
      'DOLLARFR',
      'DURATION',
      'EFFECT',
      'FV',
      'FVSCHEDULE',
      'INTRATE',
      'IPMT',
      'IRR',
      'ISPMT',
      'MDURATION',
      'MIRR',
      'NOMINAL',
      'NPER',
      'NPV',
      'ODDFPRICE',
      'ODDFYIELD',
      'ODDLPRICE',
      'ODDLYIELD',
      'PDURATION',
      'PMT',
      'PPMT',
      'PRICE',
      'PRICEDISC',
      'PRICEMAT',
      'PV',
      'RATE',
      'RECEIVED',
      'RRI',
      'SLN',
      'SYD',
      'TBILLEQ',
      'TBILLPRICE',
      'TBILLYIELD',
      'VDB',
      'XIRR',
      'XNPV',
      'YIELD',
      'YIELDDISC',
      'YIELDMAT',
    ],
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
      'AVEDEV',
      'AVERAGE',
      'AVERAGEA',
      'AVERAGEIF',
      'AVERAGEIFS',
      'BETA.DIST',
      'BETA.INV',
      'BINOM.DIST',
      'BINOM.DIST.RANGE',
      'BINOM.INV',
      'CHISQ.DIST',
      'CHISQ.DIST.RT',
      'CHISQ.INV',
      'CHISQ.INV.RT',
      'CHISQ.TEST',
      'CONFIDENCE.NORM',
      'CONFIDENCE.T',
      'CORREL',
      'COUNT',
      'COUNTA',
      'COUNTBLANK',
      'COUNTIF',
      'COUNTIFS',
      'COVARIANCE.P',
      'COVARIANCE.S',
      'DEVSQ',
      'EXPON.DIST',
      'F.DIST',
      'F.DIST.RT',
      'F.INV',
      'F.INV.RT',
      'F.TEST',
      'FISHER',
      'FISHERINV',
      'FORECAST',
      'FORECAST.ETS',
      'FORECAST.ETS.CONFINT',
      'FORECAST.ETS.SEASONALITY',
      'FORECAST.ETS.STAT',
      'FORECAST.LINEAR',
      'FREQUENCY',
      'GAMMA',
      'GAMMA.DIST',
      'GAMMA.INV',
      'GAMMALN',
      'GAMMALN.PRECISE',
      'GAUSS',
      'GEOMEAN',
      'GROWTH',
      'HARMEAN',
      'HYPGEOM.DIST',
      'INTERCEPT',
      'KURT',
      'LARGE',
      'LINEST',
      'LOGEST',
      'LOGNORM.DIST',
      'LOGNORM.INV',
      'MAX',
      'MAXA',
      'MAXIFS',
      'MEDIAN',
      'MIN',
      'MINA',
      'MINIFS',
      'MODE.MULT',
      'MODE.SNGL',
      'NEGBINOM.DIST',
      'NORM.DIST',
      'NORM.INV',
      'NORM.S.DIST',
      'NORM.S.INV',
      'PEARSON',
      'PERCENTILE.EXC',
      'PERCENTILE.INC',
      'PERCENTRANK.EXC',
      'PERCENTRANK.INC',
      'PERMUT',
      'PERMUTATIONA',
      'PHI',
      'POISSON.DIST',
      'PROB',
      'QUARTILE.EXC',
      'QUARTILE.INC',
      'RANK.AVG',
      'RANK.EQ',
      'RSQ',
      'SKEW',
      'SKEW.P',
      'SLOPE',
      'SMALL',
      'STANDARDIZE',
      'STDEV.P',
      'STDEV.S',
      'STDEVA',
      'STDEVPA',
      'STEYX',
      'T.DIST',
      'T.DIST.2T',
      'T.DIST.RT',
      'T.INV',
      'T.INV.2T',
      'T.TEST',
      'TREND',
      'TRIMMEAN',
      'VAR.P',
      'VAR.S',
      'VARA',
      'VARPA',
      'WEIBULL.DIST',
      'Z.TEST',
    ],
    engineering: [
      'BESSELI',
      'BESSELJ',
      'BESSELK',
      'BESSELY',
      'BIN2DEC',
      'BIN2HEX',
      'BIN2OCT',
      'BITAND',
      'BITLSHIFT',
      'BITOR',
      'BITRSHIFT',
      'BITXOR',
      'COMPLEX',
      'CONVERT',
      'DEC2BIN',
      'DEC2HEX',
      'DEC2OCT',
      'DELTA',
      'ERF',
      'ERF.PRECISE',
      'ERFC',
      'ERFC.PRECISE',
      'GESTEP',
      'HEX2BIN',
      'HEX2DEC',
      'HEX2OCT',
      'IMABS',
      'IMAGINARY',
      'IMARGUMENT',
      'IMCONJUGATE',
      'IMCOS',
      'IMCOSH',
      'IMCOT',
      'IMCSC',
      'IMCSCH',
      'IMDIV',
      'IMEXP',
      'IMLN',
      'IMLOG10',
      'IMLOG2',
      'IMPOWER',
      'IMPRODUCT',
      'IMREAL',
      'IMSEC',
      'IMSECH',
      'IMSIN',
      'IMSINH',
      'IMSQRT',
      'IMSUB',
      'IMSUM',
      'IMTAN',
      'OCT2BIN',
      'OCT2DEC',
      'OCT2HEX',
    ],
    information: [
      'CELL',
      'ERROR.TYPE',
      'INFO',
      'ISBLANK',
      'ISERR',
      'ISERROR',
      'ISEVEN',
      'ISFORMULA',
      'ISLOGICAL',
      'ISNA',
      'ISNONTEXT',
      'ISNUMBER',
      'ISODD',
      'ISOMITTED',
      'ISREF',
      'ISTEXT',
      'N',
      'NA',
      'SHEET',
      'SHEETS',
      'STOCKHISTORY',
      'TYPE',
    ],
    database: [
      'DAVERAGE',
      'DCOUNT',
      'DCOUNTA',
      'DGET',
      'DMAX',
      'DMIN',
      'DPRODUCT',
      'DSTDEV',
      'DSTDEVP',
      'DSUM',
      'DVAR',
      'DVARP',
    ],
    compatibility: [
      'BETADIST',
      'BETAINV',
      'BINOMDIST',
      'CHIDIST',
      'CHIINV',
      'CHITEST',
      'CONFIDENCE',
      'COVAR',
      'CRITBINOM',
      'EXPONDIST',
      'FDIST',
      'FINV',
      'FLOOR',
      'FTEST',
      'GAMMADIST',
      'GAMMAINV',
      'HYPGEOMDIST',
      'LOGINV',
      'LOGNORMDIST',
      'MODE',
      'NEGBINOMDIST',
      'NORMDIST',
      'NORMINV',
      'NORMSDIST',
      'NORMSINV',
      'PERCENTILE',
      'PERCENTRANK',
      'POISSON',
      'QUARTILE',
      'RANK',
      'STDEV',
      'STDEVP',
      'TDIST',
      'TINV',
      'TTEST',
      'VAR',
      'VARP',
      'WEIBULL',
      'ZTEST',
    ],
    cube: [
      'CUBEKPIMEMBER',
      'CUBEMEMBER',
      'CUBEMEMBERPROPERTY',
      'CUBERANKEDMEMBER',
      'CUBESET',
      'CUBESETCOUNT',
      'CUBEVALUE',
    ],
    web: ['ENCODEURL', 'FILTERXML', 'WEBSERVICE'],
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
      ...(source === 'engine' && metadata?.availability !== undefined
        ? { availability: metadata.availability }
        : {}),
    });
  }
  return {
    names: Object.freeze(uniqueNames),
    knownNames: new Set(uniqueNames),
    entries,
    source,
  };
};
