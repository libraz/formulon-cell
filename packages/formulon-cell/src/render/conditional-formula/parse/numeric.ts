import { splitFormulaArgs, splitFormulaArgsAllowEmpty } from '../splitter.js';
import type { FormulaOperand } from '../types.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type NumericCallParserContext = Pick<FormulaParserContext, 'parseFormulaOperand'>;

/** Calls that evaluate through `numericFunction` (math, engineering, probability
 *  distributions, scalar financial functions), with their arity and argument
 *  defaults, plus ISEVEN/ISODD. */
export function createNumericCallParser(ctx: NumericCallParserContext): FormulaCallParser {
  const { parseFormulaOperand } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (
      fn === 'ABS' ||
      fn === 'MOD' ||
      fn === 'ROUND' ||
      fn === 'ROUNDUP' ||
      fn === 'ROUNDDOWN' ||
      fn === 'MROUND' ||
      fn === 'QUOTIENT' ||
      fn === 'INT' ||
      fn === 'TRUNC' ||
      fn === 'SQRT' ||
      fn === 'POWER' ||
      fn === 'PI' ||
      fn === 'RADIANS' ||
      fn === 'DEGREES' ||
      fn === 'SIN' ||
      fn === 'COS' ||
      fn === 'TAN' ||
      fn === 'SEC' ||
      fn === 'CSC' ||
      fn === 'COT' ||
      fn === 'ASIN' ||
      fn === 'ACOS' ||
      fn === 'ATAN' ||
      fn === 'ATAN2' ||
      fn === 'ACOT' ||
      fn === 'SINH' ||
      fn === 'COSH' ||
      fn === 'TANH' ||
      fn === 'COTH' ||
      fn === 'SECH' ||
      fn === 'CSCH' ||
      fn === 'ASINH' ||
      fn === 'ACOSH' ||
      fn === 'ATANH' ||
      fn === 'ACOTH' ||
      fn === 'EXP' ||
      fn === 'LN' ||
      fn === 'LOG' ||
      fn === 'LOG10' ||
      fn === 'CHAR' ||
      fn === 'CODE' ||
      fn === 'UNICHAR' ||
      fn === 'UNICODE' ||
      fn === 'ADDRESS' ||
      fn === 'TYPE' ||
      fn === 'ERROR.TYPE' ||
      fn === 'FISHER' ||
      fn === 'FISHERINV' ||
      fn === 'ERF' ||
      fn === 'ERF.PRECISE' ||
      fn === 'ERFC' ||
      fn === 'ERFC.PRECISE' ||
      fn === 'GAUSS' ||
      fn === 'BASE' ||
      fn === 'DECIMAL' ||
      fn === 'BIN2DEC' ||
      fn === 'DEC2BIN' ||
      fn === 'HEX2DEC' ||
      fn === 'DEC2HEX' ||
      fn === 'OCT2DEC' ||
      fn === 'DEC2OCT' ||
      fn === 'BIN2HEX' ||
      fn === 'HEX2BIN' ||
      fn === 'BIN2OCT' ||
      fn === 'OCT2BIN' ||
      fn === 'HEX2OCT' ||
      fn === 'OCT2HEX' ||
      fn === 'ROMAN' ||
      fn === 'ARABIC' ||
      fn === 'DELTA' ||
      fn === 'GESTEP' ||
      fn === 'BITAND' ||
      fn === 'BITOR' ||
      fn === 'BITXOR' ||
      fn === 'BITLSHIFT' ||
      fn === 'BITRSHIFT' ||
      fn === 'SQRTPI' ||
      fn === 'SUMSQ' ||
      fn === 'SIGN' ||
      fn === 'GAMMA' ||
      fn === 'GAMMALN' ||
      fn === 'GAMMALN.PRECISE' ||
      fn === 'GCD' ||
      fn === 'LCM' ||
      fn === 'FACT' ||
      fn === 'FACTDOUBLE' ||
      fn === 'COMBIN' ||
      fn === 'COMBINA' ||
      fn === 'PERMUT' ||
      fn === 'PERMUTATIONA' ||
      fn === 'MULTINOMIAL' ||
      fn === 'EVEN' ||
      fn === 'ODD' ||
      fn === 'STANDARDIZE' ||
      fn === 'PHI' ||
      fn === 'CONFIDENCE' ||
      fn === 'CONFIDENCE.NORM' ||
      fn === 'CONFIDENCE.T' ||
      fn === 'PMT' ||
      fn === 'PV' ||
      fn === 'FV' ||
      fn === 'NPER' ||
      fn === 'RATE' ||
      fn === 'IPMT' ||
      fn === 'PPMT' ||
      fn === 'CUMIPMT' ||
      fn === 'CUMPRINC' ||
      fn === 'ISPMT' ||
      fn === 'EFFECT' ||
      fn === 'NOMINAL' ||
      fn === 'DOLLARDE' ||
      fn === 'DOLLARFR' ||
      fn === 'DISC' ||
      fn === 'INTRATE' ||
      fn === 'PRICEDISC' ||
      fn === 'RECEIVED' ||
      fn === 'ACCRINTM' ||
      fn === 'TBILLPRICE' ||
      fn === 'TBILLYIELD' ||
      fn === 'TBILLEQ' ||
      fn === 'RRI' ||
      fn === 'PDURATION' ||
      fn === 'SLN' ||
      fn === 'SYD' ||
      fn === 'DDB' ||
      fn === 'DB' ||
      fn === 'NORMSDIST' ||
      fn === 'NORMDIST' ||
      fn === 'NORM.S.DIST' ||
      fn === 'NORM.DIST' ||
      fn === 'NORMSINV' ||
      fn === 'NORM.S.INV' ||
      fn === 'NORMINV' ||
      fn === 'NORM.INV' ||
      fn === 'LOGINV' ||
      fn === 'LOGNORM.INV' ||
      fn === 'LOGNORMDIST' ||
      fn === 'LOGNORM.DIST' ||
      fn === 'GAMMADIST' ||
      fn === 'GAMMA.DIST' ||
      fn === 'GAMMAINV' ||
      fn === 'GAMMA.INV' ||
      fn === 'BETADIST' ||
      fn === 'BETA.DIST' ||
      fn === 'BETAINV' ||
      fn === 'BETA.INV' ||
      fn === 'FDIST' ||
      fn === 'F.DIST' ||
      fn === 'F.DIST.RT' ||
      fn === 'FINV' ||
      fn === 'F.INV' ||
      fn === 'F.INV.RT' ||
      fn === 'TDIST' ||
      fn === 'T.DIST' ||
      fn === 'T.DIST.2T' ||
      fn === 'T.DIST.RT' ||
      fn === 'TINV' ||
      fn === 'T.INV' ||
      fn === 'T.INV.2T' ||
      fn === 'CHIDIST' ||
      fn === 'CHISQ.DIST' ||
      fn === 'CHISQ.DIST.RT' ||
      fn === 'CHIINV' ||
      fn === 'CHISQ.INV' ||
      fn === 'CHISQ.INV.RT' ||
      fn === 'WEIBULL' ||
      fn === 'WEIBULL.DIST' ||
      fn === 'BINOMDIST' ||
      fn === 'BINOM.DIST' ||
      fn === 'CRITBINOM' ||
      fn === 'BINOM.INV' ||
      fn === 'NEGBINOMDIST' ||
      fn === 'NEGBINOM.DIST' ||
      fn === 'HYPGEOMDIST' ||
      fn === 'HYPGEOM.DIST' ||
      fn === 'POISSON' ||
      fn === 'POISSON.DIST' ||
      fn === 'EXPONDIST' ||
      fn === 'EXPON.DIST' ||
      fn === 'CEILING' ||
      fn === 'FLOOR' ||
      fn === 'CEILING.MATH' ||
      fn === 'FLOOR.MATH' ||
      fn === 'CEILING.PRECISE' ||
      fn === 'FLOOR.PRECISE' ||
      fn === 'ISO.CEILING'
    ) {
      if (fn === 'PI' && (aggregate[2] ?? '').trim() === '') {
        return { kind: 'numeric-function', fn, args: [] };
      }
      const args =
        fn === 'CEILING.MATH' ||
        fn === 'CEILING.PRECISE' ||
        fn === 'FLOOR.PRECISE' ||
        fn === 'ISO.CEILING' ||
        fn === 'FLOOR.MATH' ||
        fn === 'LOG' ||
        fn === 'TRUNC' ||
        fn === 'DELTA' ||
        fn === 'GESTEP' ||
        fn === 'PMT' ||
        fn === 'PV' ||
        fn === 'FV' ||
        fn === 'NPER' ||
        fn === 'RATE' ||
        fn === 'IPMT' ||
        fn === 'PPMT' ||
        fn === 'CUMIPMT' ||
        fn === 'CUMPRINC' ||
        fn === 'DISC' ||
        fn === 'INTRATE' ||
        fn === 'PRICEDISC' ||
        fn === 'RECEIVED' ||
        fn === 'ACCRINTM' ||
        fn === 'TBILLPRICE' ||
        fn === 'TBILLYIELD' ||
        fn === 'TBILLEQ' ||
        fn === 'DDB' ||
        fn === 'DB' ||
        fn === 'ADDRESS'
          ? splitFormulaArgsAllowEmpty(aggregate[2] ?? '')
          : splitFormulaArgs(aggregate[2] ?? '');
      const validLength =
        fn === 'ABS' ||
        fn === 'INT' ||
        fn === 'SQRT' ||
        fn === 'RADIANS' ||
        fn === 'DEGREES' ||
        fn === 'SIN' ||
        fn === 'COS' ||
        fn === 'TAN' ||
        fn === 'SEC' ||
        fn === 'CSC' ||
        fn === 'COT' ||
        fn === 'ASIN' ||
        fn === 'ACOS' ||
        fn === 'ATAN' ||
        fn === 'ACOT' ||
        fn === 'SINH' ||
        fn === 'COSH' ||
        fn === 'TANH' ||
        fn === 'COTH' ||
        fn === 'SECH' ||
        fn === 'CSCH' ||
        fn === 'ASINH' ||
        fn === 'ACOSH' ||
        fn === 'ATANH' ||
        fn === 'ACOTH' ||
        fn === 'EXP' ||
        fn === 'LN' ||
        fn === 'LOG10' ||
        fn === 'CHAR' ||
        fn === 'CODE' ||
        fn === 'UNICHAR' ||
        fn === 'UNICODE' ||
        fn === 'TYPE' ||
        fn === 'ERROR.TYPE' ||
        fn === 'FISHER' ||
        fn === 'FISHERINV' ||
        fn === 'ERF.PRECISE' ||
        fn === 'ERFC' ||
        fn === 'ERFC.PRECISE' ||
        fn === 'GAUSS' ||
        fn === 'ARABIC' ||
        fn === 'SQRTPI' ||
        fn === 'SIGN' ||
        fn === 'GAMMA' ||
        fn === 'GAMMALN' ||
        fn === 'GAMMALN.PRECISE' ||
        fn === 'FACT' ||
        fn === 'FACTDOUBLE' ||
        fn === 'EVEN' ||
        fn === 'ODD' ||
        fn === 'NORMSDIST' ||
        fn === 'PHI'
          ? args?.length === 1
          : fn === 'ADDRESS'
            ? args !== null && args.length >= 2 && args.length <= 5
            : fn === 'GCD' || fn === 'LCM' || fn === 'SUMSQ' || fn === 'MULTINOMIAL'
              ? args !== null && args.length > 0
              : fn === 'LOG'
                ? args?.length === 1 || args?.length === 2
                : fn === 'ERF'
                  ? args?.length === 1 || args?.length === 2
                  : fn === 'BASE'
                    ? args?.length === 2 || args?.length === 3
                    : fn === 'DEC2BIN' ||
                        fn === 'DEC2HEX' ||
                        fn === 'DEC2OCT' ||
                        fn === 'BIN2HEX' ||
                        fn === 'HEX2BIN' ||
                        fn === 'BIN2OCT' ||
                        fn === 'OCT2BIN' ||
                        fn === 'HEX2OCT' ||
                        fn === 'OCT2HEX'
                      ? args?.length === 1 || args?.length === 2
                      : fn === 'DECIMAL'
                        ? args?.length === 2
                        : fn === 'ROMAN'
                          ? args?.length === 1 || args?.length === 2
                          : fn === 'BIN2DEC' || fn === 'HEX2DEC' || fn === 'OCT2DEC'
                            ? args?.length === 1
                            : fn === 'DELTA' || fn === 'GESTEP'
                              ? args?.length === 1 || args?.length === 2
                              : fn === 'BITAND' ||
                                  fn === 'BITOR' ||
                                  fn === 'BITXOR' ||
                                  fn === 'BITLSHIFT' ||
                                  fn === 'BITRSHIFT'
                                ? args?.length === 2
                                : fn === 'STANDARDIZE'
                                  ? args?.length === 3
                                  : fn === 'RATE'
                                    ? args?.length === 3 ||
                                      args?.length === 4 ||
                                      args?.length === 5 ||
                                      args?.length === 6
                                    : fn === 'PMT' || fn === 'PV' || fn === 'FV' || fn === 'NPER'
                                      ? args?.length === 3 ||
                                        args?.length === 4 ||
                                        args?.length === 5
                                      : fn === 'IPMT' || fn === 'PPMT'
                                        ? args?.length === 4 ||
                                          args?.length === 5 ||
                                          args?.length === 6
                                        : fn === 'CUMIPMT' || fn === 'CUMPRINC'
                                          ? args?.length === 6
                                          : fn === 'ISPMT'
                                            ? args?.length === 4
                                            : fn === 'EFFECT' || fn === 'NOMINAL'
                                              ? args?.length === 2
                                              : fn === 'DOLLARDE' || fn === 'DOLLARFR'
                                                ? args?.length === 2
                                                : fn === 'DISC' ||
                                                    fn === 'INTRATE' ||
                                                    fn === 'PRICEDISC' ||
                                                    fn === 'RECEIVED'
                                                  ? args?.length === 4 || args?.length === 5
                                                  : fn === 'ACCRINTM'
                                                    ? args?.length === 3 ||
                                                      args?.length === 4 ||
                                                      args?.length === 5
                                                    : fn === 'TBILLPRICE' ||
                                                        fn === 'TBILLYIELD' ||
                                                        fn === 'TBILLEQ'
                                                      ? args?.length === 3
                                                      : fn === 'RRI' || fn === 'PDURATION'
                                                        ? args?.length === 3
                                                        : fn === 'SLN'
                                                          ? args?.length === 3
                                                          : fn === 'SYD'
                                                            ? args?.length === 4
                                                            : fn === 'DDB'
                                                              ? args?.length === 4 ||
                                                                args?.length === 5
                                                              : fn === 'DB'
                                                                ? args?.length === 4 ||
                                                                  args?.length === 5
                                                                : fn === 'NORM.S.DIST'
                                                                  ? args?.length === 2
                                                                  : fn === 'CONFIDENCE' ||
                                                                      fn === 'CONFIDENCE.NORM' ||
                                                                      fn === 'CONFIDENCE.T'
                                                                    ? args?.length === 3
                                                                    : fn === 'NORMSINV' ||
                                                                        fn === 'NORM.S.INV'
                                                                      ? args?.length === 1
                                                                      : fn === 'NORMINV' ||
                                                                          fn === 'NORM.INV' ||
                                                                          fn === 'LOGINV' ||
                                                                          fn === 'LOGNORM.INV'
                                                                        ? args?.length === 3
                                                                        : fn === 'NORMDIST' ||
                                                                            fn === 'NORM.DIST'
                                                                          ? args?.length === 4
                                                                          : fn === 'LOGNORMDIST'
                                                                            ? args?.length === 3
                                                                            : fn ===
                                                                                  'LOGNORM.DIST' ||
                                                                                fn ===
                                                                                  'GAMMADIST' ||
                                                                                fn ===
                                                                                  'GAMMA.DIST' ||
                                                                                fn === 'WEIBULL' ||
                                                                                fn ===
                                                                                  'WEIBULL.DIST'
                                                                              ? args?.length === 4
                                                                              : fn === 'GAMMAINV' ||
                                                                                  fn === 'GAMMA.INV'
                                                                                ? args?.length === 3
                                                                                : fn ===
                                                                                      'BETADIST' ||
                                                                                    fn === 'BETAINV'
                                                                                  ? args?.length ===
                                                                                      3 ||
                                                                                    args?.length ===
                                                                                      4 ||
                                                                                    args?.length ===
                                                                                      5
                                                                                  : fn ===
                                                                                      'BETA.DIST'
                                                                                    ? args?.length ===
                                                                                        4 ||
                                                                                      args?.length ===
                                                                                        5 ||
                                                                                      args?.length ===
                                                                                        6
                                                                                    : fn ===
                                                                                        'BETA.INV'
                                                                                      ? args?.length ===
                                                                                          3 ||
                                                                                        args?.length ===
                                                                                          4 ||
                                                                                        args?.length ===
                                                                                          5
                                                                                      : fn ===
                                                                                            'FDIST' ||
                                                                                          fn ===
                                                                                            'F.DIST.RT' ||
                                                                                          fn ===
                                                                                            'FINV' ||
                                                                                          fn ===
                                                                                            'F.INV' ||
                                                                                          fn ===
                                                                                            'F.INV.RT'
                                                                                        ? args?.length ===
                                                                                          3
                                                                                        : fn ===
                                                                                            'F.DIST'
                                                                                          ? args?.length ===
                                                                                            4
                                                                                          : fn ===
                                                                                                'TDIST' ||
                                                                                              fn ===
                                                                                                'T.DIST'
                                                                                            ? args?.length ===
                                                                                              3
                                                                                            : fn ===
                                                                                                  'T.DIST.2T' ||
                                                                                                fn ===
                                                                                                  'T.DIST.RT' ||
                                                                                                fn ===
                                                                                                  'TINV' ||
                                                                                                fn ===
                                                                                                  'T.INV' ||
                                                                                                fn ===
                                                                                                  'T.INV.2T'
                                                                                              ? args?.length ===
                                                                                                2
                                                                                              : fn ===
                                                                                                    'CHIDIST' ||
                                                                                                  fn ===
                                                                                                    'CHISQ.DIST.RT' ||
                                                                                                  fn ===
                                                                                                    'CHIINV' ||
                                                                                                  fn ===
                                                                                                    'CHISQ.INV' ||
                                                                                                  fn ===
                                                                                                    'CHISQ.INV.RT'
                                                                                                ? args?.length ===
                                                                                                  2
                                                                                                : fn ===
                                                                                                    'CHISQ.DIST'
                                                                                                  ? args?.length ===
                                                                                                    3
                                                                                                  : fn ===
                                                                                                        'BINOMDIST' ||
                                                                                                      fn ===
                                                                                                        'BINOM.DIST'
                                                                                                    ? args?.length ===
                                                                                                      4
                                                                                                    : fn ===
                                                                                                          'CRITBINOM' ||
                                                                                                        fn ===
                                                                                                          'BINOM.INV'
                                                                                                      ? args?.length ===
                                                                                                        3
                                                                                                      : fn ===
                                                                                                          'NEGBINOMDIST'
                                                                                                        ? args?.length ===
                                                                                                          3
                                                                                                        : fn ===
                                                                                                            'NEGBINOM.DIST'
                                                                                                          ? args?.length ===
                                                                                                            4
                                                                                                          : fn ===
                                                                                                              'HYPGEOMDIST'
                                                                                                            ? args?.length ===
                                                                                                              4
                                                                                                            : fn ===
                                                                                                                'HYPGEOM.DIST'
                                                                                                              ? args?.length ===
                                                                                                                5
                                                                                                              : fn ===
                                                                                                                    'POISSON' ||
                                                                                                                  fn ===
                                                                                                                    'POISSON.DIST' ||
                                                                                                                  fn ===
                                                                                                                    'EXPONDIST' ||
                                                                                                                  fn ===
                                                                                                                    'EXPON.DIST'
                                                                                                                ? args?.length ===
                                                                                                                  3
                                                                                                                : fn ===
                                                                                                                      'CEILING' ||
                                                                                                                    fn ===
                                                                                                                      'FLOOR' ||
                                                                                                                    fn ===
                                                                                                                      'MROUND' ||
                                                                                                                    fn ===
                                                                                                                      'QUOTIENT'
                                                                                                                  ? args?.length ===
                                                                                                                    2
                                                                                                                  : fn ===
                                                                                                                        'COMBIN' ||
                                                                                                                      fn ===
                                                                                                                        'COMBINA' ||
                                                                                                                      fn ===
                                                                                                                        'PERMUT' ||
                                                                                                                      fn ===
                                                                                                                        'PERMUTATIONA'
                                                                                                                    ? args?.length ===
                                                                                                                      2
                                                                                                                    : fn ===
                                                                                                                          'CEILING.MATH' ||
                                                                                                                        fn ===
                                                                                                                          'FLOOR.MATH'
                                                                                                                      ? args?.length ===
                                                                                                                          1 ||
                                                                                                                        args?.length ===
                                                                                                                          2 ||
                                                                                                                        args?.length ===
                                                                                                                          3
                                                                                                                      : fn ===
                                                                                                                            'CEILING.PRECISE' ||
                                                                                                                          fn ===
                                                                                                                            'FLOOR.PRECISE' ||
                                                                                                                          fn ===
                                                                                                                            'ISO.CEILING'
                                                                                                                        ? args?.length ===
                                                                                                                            1 ||
                                                                                                                          args?.length ===
                                                                                                                            2
                                                                                                                        : fn ===
                                                                                                                            'TRUNC'
                                                                                                                          ? args?.length ===
                                                                                                                              1 ||
                                                                                                                            args?.length ===
                                                                                                                              2
                                                                                                                          : args?.length ===
                                                                                                                            2;
      if (args && validLength) {
        if (
          (fn === 'CEILING.MATH' ||
            fn === 'CEILING.PRECISE' ||
            fn === 'FLOOR.PRECISE' ||
            fn === 'ISO.CEILING' ||
            fn === 'FLOOR.MATH' ||
            fn === 'LOG' ||
            fn === 'TRUNC' ||
            fn === 'DELTA' ||
            fn === 'GESTEP' ||
            fn === 'PMT' ||
            fn === 'PV' ||
            fn === 'FV' ||
            fn === 'NPER' ||
            fn === 'RATE' ||
            fn === 'IPMT' ||
            fn === 'PPMT' ||
            fn === 'CUMIPMT' ||
            fn === 'CUMPRINC' ||
            fn === 'DISC' ||
            fn === 'INTRATE' ||
            fn === 'PRICEDISC' ||
            fn === 'RECEIVED' ||
            fn === 'ACCRINTM' ||
            fn === 'TBILLPRICE' ||
            fn === 'TBILLYIELD' ||
            fn === 'TBILLEQ' ||
            fn === 'DDB' ||
            fn === 'DB' ||
            fn === 'ADDRESS') &&
          (args[0] ?? '').trim() === ''
        ) {
          return null;
        }
        const operands = args.map((arg, index) => {
          if (
            (fn === 'CEILING.MATH' ||
              fn === 'FLOOR.MATH' ||
              fn === 'CEILING.PRECISE' ||
              fn === 'FLOOR.PRECISE' ||
              fn === 'ISO.CEILING') &&
            arg.trim() === ''
          ) {
            const value =
              fn === 'CEILING.PRECISE' || fn === 'FLOOR.PRECISE' || fn === 'ISO.CEILING'
                ? 1
                : index === 1
                  ? 1
                  : 0;
            return { kind: 'literal' as const, value: { kind: 'number' as const, value } };
          }
          if (fn === 'LOG' && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 10 } };
          }
          if (fn === 'TRUNC' && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if ((fn === 'DELTA' || fn === 'GESTEP') && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if (fn === 'ADDRESS' && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } };
          }
          if (fn === 'ADDRESS' && index === 3 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'bool' as const, value: true } };
          }
          if (fn === 'ADDRESS' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'text' as const, value: '' } };
          }
          if (
            arg.trim() === '' &&
            (((fn === 'PMT' || fn === 'PV' || fn === 'FV' || fn === 'NPER' || fn === 'RATE') &&
              index >= 3) ||
              ((fn === 'IPMT' || fn === 'PPMT') && index >= 4))
          ) {
            return {
              kind: 'literal' as const,
              value: { kind: 'number' as const, value: fn === 'RATE' && index === 5 ? 0.1 : 0 },
            };
          }
          if (fn === 'DDB' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 2 } };
          }
          if (fn === 'DB' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 12 } };
          }
          if (
            (fn === 'DISC' || fn === 'INTRATE' || fn === 'PRICEDISC' || fn === 'RECEIVED') &&
            index === 4 &&
            arg.trim() === ''
          ) {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if (fn === 'ACCRINTM' && index === 3 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1000 } };
          }
          if (fn === 'ACCRINTM' && index === 4 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          return parseFormulaOperand(arg, sheetIndex);
        });
        if (operands.every((operand) => operand !== null)) {
          return { kind: 'numeric-function', fn, args: operands as FormulaOperand[] };
        }
      }
    }
    if (fn === 'ISEVEN' || fn === 'ISODD') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'numeric-predicate', fn, value };
      }
    }
    return undefined;
  };
}
