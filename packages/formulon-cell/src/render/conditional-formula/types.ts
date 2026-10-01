import type { CellValue } from '../../engine/types.js';

export interface ParsedRef {
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
}

export interface ParsedA1Range {
  start: ParsedRef;
  end: ParsedRef;
}

export type FormulaAggregateName =
  | 'SUM'
  | 'AVERAGE'
  | 'AVERAGEA'
  | 'MIN'
  | 'MINA'
  | 'MAX'
  | 'MAXA'
  | 'COUNT'
  | 'COUNTA'
  | 'COUNTBLANK'
  | 'PRODUCT'
  | 'MEDIAN'
  | 'MODE'
  | 'MODE.SNGL'
  | 'AVEDEV'
  | 'DEVSQ'
  | 'SKEW'
  | 'SKEW.P'
  | 'KURT'
  | 'GEOMEAN'
  | 'HARMEAN'
  | 'STDEV'
  | 'STDEVP'
  | 'STDEV.S'
  | 'STDEV.P'
  | 'VAR'
  | 'VARP'
  | 'VAR.S'
  | 'VAR.P';

export type FormulaRangeArg =
  | { kind: 'range'; range: ParsedA1Range }
  | { kind: 'dynamic-range'; range: FormulaRangeOperand };

export type FormulaAggregateArg = FormulaRangeArg | { kind: 'operand'; operand: FormulaOperand };

export type FormulaRangeOperand =
  | {
      kind: 'offset-range';
      reference: FormulaRangeArg;
      rows: FormulaOperand;
      cols: FormulaOperand;
      height?: FormulaOperand;
      width?: FormulaOperand;
    }
  | { kind: 'indirect-range'; refText: FormulaOperand; a1?: FormulaOperand };

export type FormulaOperand =
  | { kind: 'ref'; ref: ParsedRef }
  | { kind: 'range-aggregate'; fn: FormulaAggregateName; range: FormulaRangeArg }
  | { kind: 'aggregate-args'; fn: FormulaAggregateName; args: FormulaAggregateArg[] }
  | { kind: 'subtotal'; functionNum: FormulaOperand; args: FormulaAggregateArg[] }
  | {
      kind: 'aggregate-function';
      functionNum: FormulaOperand;
      options: FormulaOperand;
      args: FormulaAggregateArg[];
    }
  | { kind: 'ranked-range'; fn: 'LARGE' | 'SMALL'; range: FormulaRangeArg; rank: FormulaOperand }
  | {
      kind: 'percentile-range';
      fn:
        | 'PERCENTILE.INC'
        | 'PERCENTILE.EXC'
        | 'PERCENTILE'
        | 'QUARTILE.INC'
        | 'QUARTILE.EXC'
        | 'QUARTILE'
        | 'PERCENTRANK'
        | 'PERCENTRANK.INC'
        | 'PERCENTRANK.EXC';
      range: FormulaRangeArg;
      value: FormulaOperand;
      significance?: FormulaOperand;
    }
  | {
      kind: 'range-rank';
      fn: 'RANK' | 'RANK.EQ' | 'RANK.AVG';
      value: FormulaOperand;
      range: FormulaRangeArg;
      order?: FormulaOperand;
    }
  | {
      kind: 'paired-range-stat';
      fn:
        | 'CORREL'
        | 'PEARSON'
        | 'COVAR'
        | 'COVARIANCE.P'
        | 'COVARIANCE.S'
        | 'SLOPE'
        | 'INTERCEPT'
        | 'RSQ'
        | 'STEYX'
        | 'SUMX2MY2'
        | 'SUMX2PY2'
        | 'SUMXMY2'
        | 'F.TEST'
        | 'FTEST';
      left: FormulaRangeArg;
      right: FormulaRangeArg;
    }
  | {
      kind: 'regression-forecast';
      fn: 'FORECAST' | 'FORECAST.LINEAR';
      x: FormulaOperand;
      knownY: FormulaRangeArg;
      knownX: FormulaRangeArg;
    }
  | {
      kind: 'probability-range';
      values: FormulaRangeArg;
      probabilities: FormulaRangeArg;
      lower: FormulaOperand;
      upper?: FormulaOperand;
    }
  | {
      kind: 'z-test';
      range: FormulaRangeArg;
      x: FormulaOperand;
      sigma?: FormulaOperand;
    }
  | {
      kind: 't-test';
      left: FormulaRangeArg;
      right: FormulaRangeArg;
      tails: FormulaOperand;
      type: FormulaOperand;
    }
  | {
      kind: 'chisq-test';
      actual: FormulaRangeArg;
      expected: FormulaRangeArg;
    }
  | {
      kind: 'series-sum';
      x: FormulaOperand;
      n: FormulaOperand;
      m: FormulaOperand;
      coefficients: FormulaAggregateArg[];
    }
  | { kind: 'npv'; rate: FormulaOperand; values: FormulaAggregateArg[] }
  | {
      kind: 'mirr';
      values: FormulaRangeArg;
      financeRate: FormulaOperand;
      reinvestRate: FormulaOperand;
    }
  | {
      kind: 'xnpv';
      rate: FormulaOperand;
      values: FormulaRangeArg;
      dates: FormulaRangeArg;
    }
  | {
      kind: 'xirr';
      values: FormulaRangeArg;
      dates: FormulaRangeArg;
      guess?: FormulaOperand;
    }
  | { kind: 'irr'; values: FormulaRangeArg; guess?: FormulaOperand }
  | { kind: 'fv-schedule'; principal: FormulaOperand; schedule: FormulaRangeArg }
  | { kind: 'sumproduct'; ranges: FormulaRangeArg[] }
  | { kind: 'countif'; range: FormulaRangeArg; criteria: FormulaOperand }
  | { kind: 'countifs'; pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] }
  | { kind: 'sumif'; range: FormulaRangeArg; criteria: FormulaOperand; sumRange: FormulaRangeArg }
  | {
      kind: 'averageif';
      range: FormulaRangeArg;
      criteria: FormulaOperand;
      averageRange: FormulaRangeArg;
    }
  | {
      kind: 'sumifs';
      sumRange: FormulaRangeArg;
      pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[];
    }
  | {
      kind: 'averageifs';
      averageRange: FormulaRangeArg;
      pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[];
    }
  | {
      kind: 'minmaxifs';
      fn: 'MINIFS' | 'MAXIFS';
      valueRange: FormulaRangeArg;
      pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[];
    }
  | { kind: 'text-length'; value: FormulaOperand }
  | { kind: 'formula-text'; ref: FormulaRangeArg }
  | {
      kind: 'text-search';
      fn: 'SEARCH' | 'FIND';
      needle: FormulaOperand;
      haystack: FormulaOperand;
      start?: FormulaOperand;
    }
  | {
      kind: 'text-slice';
      fn: 'LEFT' | 'RIGHT' | 'MID';
      value: FormulaOperand;
      start?: FormulaOperand;
      count: FormulaOperand;
    }
  | { kind: 'text-concat-function'; values: FormulaOperand[] }
  | {
      kind: 'text-substitute';
      value: FormulaOperand;
      oldText: FormulaOperand;
      newText: FormulaOperand;
      instance?: FormulaOperand;
    }
  | {
      kind: 'text-replace';
      value: FormulaOperand;
      start: FormulaOperand;
      count: FormulaOperand;
      newText: FormulaOperand;
    }
  | { kind: 'text-repeat'; value: FormulaOperand; count: FormulaOperand }
  | {
      kind: 'text-before-after';
      fn: 'TEXTBEFORE' | 'TEXTAFTER';
      value: FormulaOperand;
      delimiter: FormulaOperand;
      instance?: FormulaOperand;
      matchMode?: FormulaOperand;
      matchEnd?: FormulaOperand;
      ifNotFound?: FormulaOperand;
    }
  | {
      kind: 'text-join';
      delimiter: FormulaOperand;
      ignoreEmpty: FormulaOperand;
      values: FormulaOperand[];
    }
  | {
      kind: 'text-transform';
      fn: 'LOWER' | 'UPPER' | 'TRIM' | 'CLEAN' | 'PROPER' | 'ENCODEURL';
      value: FormulaOperand;
    }
  | { kind: 'text-exact'; left: FormulaOperand; right: FormulaOperand }
  | { kind: 'text-format'; value: FormulaOperand; pattern: FormulaOperand }
  | {
      kind: 'text-fixed-format';
      fn: 'DOLLAR' | 'FIXED';
      value: FormulaOperand;
      decimals?: FormulaOperand;
      noCommas?: FormulaOperand;
    }
  | { kind: 'text-value'; value: FormulaOperand }
  | {
      kind: 'text-number-value';
      value: FormulaOperand;
      decimalSeparator?: FormulaOperand;
      groupSeparator?: FormulaOperand;
    }
  | { kind: 'value-to-text'; value: FormulaOperand; format?: FormulaOperand }
  | { kind: 'hyperlink'; link: FormulaOperand; friendlyName?: FormulaOperand }
  | { kind: 'scalar-coerce'; fn: 'N' | 'T'; value: FormulaOperand }
  | { kind: 'position'; fn: 'ROW' | 'COLUMN'; ref?: FormulaRangeArg }
  | { kind: 'range-dimension'; fn: 'ROWS' | 'COLUMNS' | 'AREAS'; range: FormulaRangeArg }
  | {
      kind: 'numeric-function';
      fn:
        | 'ABS'
        | 'MOD'
        | 'ROUND'
        | 'ROUNDUP'
        | 'ROUNDDOWN'
        | 'MROUND'
        | 'QUOTIENT'
        | 'INT'
        | 'TRUNC'
        | 'SQRT'
        | 'POWER'
        | 'PI'
        | 'RADIANS'
        | 'DEGREES'
        | 'SIN'
        | 'COS'
        | 'TAN'
        | 'SEC'
        | 'CSC'
        | 'COT'
        | 'ASIN'
        | 'ACOS'
        | 'ATAN'
        | 'ATAN2'
        | 'ACOT'
        | 'SINH'
        | 'COSH'
        | 'TANH'
        | 'COTH'
        | 'SECH'
        | 'CSCH'
        | 'ASINH'
        | 'ACOSH'
        | 'ATANH'
        | 'ACOTH'
        | 'EXP'
        | 'LN'
        | 'LOG'
        | 'LOG10'
        | 'CHAR'
        | 'CODE'
        | 'UNICHAR'
        | 'UNICODE'
        | 'ADDRESS'
        | 'TYPE'
        | 'ERROR.TYPE'
        | 'FISHER'
        | 'FISHERINV'
        | 'ERF'
        | 'ERF.PRECISE'
        | 'ERFC'
        | 'ERFC.PRECISE'
        | 'GAUSS'
        | 'BASE'
        | 'DECIMAL'
        | 'BIN2DEC'
        | 'DEC2BIN'
        | 'HEX2DEC'
        | 'DEC2HEX'
        | 'OCT2DEC'
        | 'DEC2OCT'
        | 'BIN2HEX'
        | 'HEX2BIN'
        | 'BIN2OCT'
        | 'OCT2BIN'
        | 'HEX2OCT'
        | 'OCT2HEX'
        | 'ROMAN'
        | 'ARABIC'
        | 'DELTA'
        | 'GESTEP'
        | 'BITAND'
        | 'BITOR'
        | 'BITXOR'
        | 'BITLSHIFT'
        | 'BITRSHIFT'
        | 'SQRTPI'
        | 'SUMSQ'
        | 'SIGN'
        | 'GAMMA'
        | 'GAMMALN'
        | 'GAMMALN.PRECISE'
        | 'GCD'
        | 'LCM'
        | 'FACT'
        | 'FACTDOUBLE'
        | 'COMBIN'
        | 'COMBINA'
        | 'PERMUT'
        | 'PERMUTATIONA'
        | 'MULTINOMIAL'
        | 'EVEN'
        | 'ODD'
        | 'STANDARDIZE'
        | 'PHI'
        | 'CONFIDENCE'
        | 'CONFIDENCE.NORM'
        | 'CONFIDENCE.T'
        | 'PMT'
        | 'PV'
        | 'FV'
        | 'NPER'
        | 'RATE'
        | 'IPMT'
        | 'PPMT'
        | 'CUMIPMT'
        | 'CUMPRINC'
        | 'ISPMT'
        | 'EFFECT'
        | 'NOMINAL'
        | 'DOLLARDE'
        | 'DOLLARFR'
        | 'DISC'
        | 'INTRATE'
        | 'PRICEDISC'
        | 'RECEIVED'
        | 'ACCRINTM'
        | 'TBILLPRICE'
        | 'TBILLYIELD'
        | 'TBILLEQ'
        | 'RRI'
        | 'PDURATION'
        | 'SLN'
        | 'SYD'
        | 'DDB'
        | 'DB'
        | 'NORMSDIST'
        | 'NORMDIST'
        | 'NORM.S.DIST'
        | 'NORM.DIST'
        | 'NORMSINV'
        | 'NORM.S.INV'
        | 'NORMINV'
        | 'NORM.INV'
        | 'LOGINV'
        | 'LOGNORM.INV'
        | 'LOGNORMDIST'
        | 'LOGNORM.DIST'
        | 'GAMMADIST'
        | 'GAMMA.DIST'
        | 'GAMMAINV'
        | 'GAMMA.INV'
        | 'BETADIST'
        | 'BETA.DIST'
        | 'BETAINV'
        | 'BETA.INV'
        | 'FDIST'
        | 'F.DIST'
        | 'F.DIST.RT'
        | 'FINV'
        | 'F.INV'
        | 'F.INV.RT'
        | 'TDIST'
        | 'T.DIST'
        | 'T.DIST.2T'
        | 'T.DIST.RT'
        | 'TINV'
        | 'T.INV'
        | 'T.INV.2T'
        | 'CHIDIST'
        | 'CHISQ.DIST'
        | 'CHISQ.DIST.RT'
        | 'CHIINV'
        | 'CHISQ.INV'
        | 'CHISQ.INV.RT'
        | 'WEIBULL'
        | 'WEIBULL.DIST'
        | 'BINOMDIST'
        | 'BINOM.DIST'
        | 'CRITBINOM'
        | 'BINOM.INV'
        | 'NEGBINOMDIST'
        | 'NEGBINOM.DIST'
        | 'HYPGEOMDIST'
        | 'HYPGEOM.DIST'
        | 'POISSON'
        | 'POISSON.DIST'
        | 'EXPONDIST'
        | 'EXPON.DIST'
        | 'CEILING'
        | 'FLOOR'
        | 'CEILING.MATH'
        | 'FLOOR.MATH'
        | 'CEILING.PRECISE'
        | 'FLOOR.PRECISE'
        | 'ISO.CEILING';
      args: FormulaOperand[];
    }
  | { kind: 'numeric-predicate'; fn: 'ISEVEN' | 'ISODD'; value: FormulaOperand }
  | {
      kind: 'date-function';
      fn:
        | 'DATE'
        | 'YEAR'
        | 'MONTH'
        | 'DAY'
        | 'WEEKDAY'
        | 'WEEKNUM'
        | 'ISOWEEKNUM'
        | 'TODAY'
        | 'NOW'
        | 'TIME'
        | 'EDATE'
        | 'EOMONTH'
        | 'DAYS'
        | 'DAYS360'
        | 'DATEDIF'
        | 'YEARFRAC'
        | 'DATEVALUE'
        | 'TIMEVALUE'
        | 'NETWORKDAYS'
        | 'NETWORKDAYS.INTL'
        | 'WORKDAY'
        | 'WORKDAY.INTL'
        | 'HOUR'
        | 'MINUTE'
        | 'SECOND';
      args: FormulaDateArg[];
    }
  | {
      kind: 'error-fallback';
      fn: 'IFERROR' | 'IFNA';
      value: FormulaOperand;
      fallback: FormulaOperand;
    }
  | { kind: 'match'; lookup: FormulaOperand; range: FormulaRangeArg; matchType?: FormulaOperand }
  | {
      kind: 'offset';
      reference: FormulaRangeArg;
      rows: FormulaOperand;
      cols: FormulaOperand;
      height?: FormulaOperand;
      width?: FormulaOperand;
    }
  | { kind: 'indirect'; refText: FormulaOperand; a1?: FormulaOperand }
  | {
      kind: 'index';
      range: FormulaRangeArg;
      row: FormulaOperand;
      col?: FormulaOperand;
    }
  | {
      kind: 'lookup';
      fn: 'VLOOKUP' | 'HLOOKUP';
      lookup: FormulaOperand;
      range: FormulaRangeArg;
      index: FormulaOperand;
      rangeLookup: FormulaOperand;
    }
  | {
      kind: 'xlookup';
      lookup: FormulaOperand;
      lookupRange: FormulaRangeArg;
      returnRange: FormulaRangeArg;
      ifNotFound?: FormulaOperand;
      matchMode?: FormulaOperand;
      searchMode?: FormulaOperand;
    }
  | {
      kind: 'xmatch';
      lookup: FormulaOperand;
      range: FormulaRangeArg;
      matchMode?: FormulaOperand;
      searchMode?: FormulaOperand;
    }
  | {
      kind: 'vector-lookup';
      lookup: FormulaOperand;
      lookupRange: FormulaRangeArg;
      resultRange?: FormulaRangeArg;
    }
  | { kind: 'cell-info'; infoType: FormulaOperand; ref?: FormulaRangeArg }
  | { kind: 'sheet-info'; fn: 'SHEET' | 'SHEETS'; range?: FormulaRangeArg }
  | { kind: 'choose'; index: FormulaOperand; choices: FormulaOperand[] }
  | {
      kind: 'switch';
      value: FormulaOperand;
      cases: { match: FormulaOperand; result: FormulaOperand }[];
      defaultValue?: FormulaOperand;
    }
  | { kind: 'if'; condition: FormulaCondition; whenTrue: FormulaOperand; whenFalse: FormulaOperand }
  | { kind: 'ifs'; branches: { condition: FormulaCondition; result: FormulaOperand }[] }
  | { kind: 'condition-value'; condition: FormulaCondition }
  | { kind: 'literal'; value: CellValue }
  | {
      kind: 'binary';
      op: '+' | '-' | '*' | '/' | '^' | '&';
      left: FormulaOperand;
      right: FormulaOperand;
    };

export type FormulaDateArg = FormulaOperand | FormulaRangeArg;

export type FormulaCondition =
  | { kind: 'bool'; value: boolean }
  | {
      kind: 'logical';
      fn: 'AND' | 'OR' | 'NOT' | 'XOR';
      args: FormulaCondition[];
    }
  | {
      kind: 'is';
      fn:
        | 'ISBLANK'
        | 'ISERROR'
        | 'ISERR'
        | 'ISNA'
        | 'ISNUMBER'
        | 'ISTEXT'
        | 'ISLOGICAL'
        | 'ISNONTEXT'
        | 'ISFORMULA'
        | 'ISREF';
      value: FormulaOperand | FormulaRangeArg;
    }
  | {
      kind: 'comparison';
      left: FormulaOperand;
      op: '>' | '<' | '>=' | '<=' | '=' | '<>';
      right: FormulaOperand;
    }
  | { kind: 'operand'; value: FormulaOperand };
