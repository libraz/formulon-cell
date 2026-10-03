import type { CellValue } from '../../../engine/types.js';
import { numericResult, readLogical, readNumber } from '../coercion.js';
import {
  binomialProbability,
  hypergeometricProbability,
  inverseRegularizedBeta,
  inverseRegularizedGammaP,
  inverseStandardNormal,
  inverseStudentTCdf,
  logGamma,
  negativeBinomialProbability,
  poissonProbability,
  regularizedBeta,
  regularizedGammaP,
  standardNormalCdf,
  standardNormalPdf,
  studentTCdf,
  studentTPdf,
} from '../distributions.js';
import type { FormulaOperand } from '../types.js';
import type { FormulaReaderContext, NumericFunctionName } from './context.js';

/** Exactly the reader members this family touches. */
export type ProbabilityEvaluatorContext = Pick<FormulaReaderContext, 'readOperand'>;

/** Probability distributions and their inverses (normal, lognormal, gamma,
 *  beta, F, t, chi-squared, binomial family, Poisson, exponential, Weibull). */
export function createProbabilityEvaluator(ctx: ProbabilityEvaluatorContext) {
  const { readOperand } = ctx;
  const probabilityFunction = (
    fn: NumericFunctionName,
    args: FormulaOperand[],
    rowOffset: number,
    colOffset: number,
  ): CellValue | undefined => {
    if (fn === 'NORMSDIST') {
      const [zOperand] = args;
      if (!zOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const z = readNumber(readOperand(zOperand, rowOffset, colOffset));
      if (z === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return numericResult(standardNormalCdf(z));
    }
    if (fn === 'NORM.S.DIST') {
      const [zOperand, cumulativeOperand] = args;
      if (!zOperand || !cumulativeOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const z = readNumber(readOperand(zOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (z === null || cumulative === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      return numericResult(cumulative ? standardNormalCdf(z) : standardNormalPdf(z));
    }
    if (fn === 'CONFIDENCE' || fn === 'CONFIDENCE.NORM' || fn === 'CONFIDENCE.T') {
      const [alphaOperand, standardDeviationOperand, sizeOperand] = args;
      if (!alphaOperand || !standardDeviationOperand || !sizeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      const size = readNumber(readOperand(sizeOperand, rowOffset, colOffset));
      if (alpha === null || standardDeviation === null || size === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const sampleSize = Math.trunc(size);
      if (alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || sampleSize < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'CONFIDENCE.T') {
        if (sampleSize < 2) return { kind: 'error', code: 6, text: '#NUM!' };
        const critical = inverseStudentTCdf(1 - alpha / 2, sampleSize - 1);
        return critical === null
          ? { kind: 'error', code: 6, text: '#NUM!' }
          : numericResult((critical * standardDeviation) / Math.sqrt(sampleSize));
      }
      return numericResult(
        (inverseStandardNormal(1 - alpha / 2) * standardDeviation) / Math.sqrt(sampleSize),
      );
    }
    if (fn === 'NORMDIST' || fn === 'NORM.DIST') {
      const [xOperand, meanOperand, standardDeviationOperand, cumulativeOperand] = args;
      if (!xOperand || !meanOperand || !standardDeviationOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || mean === null || standardDeviation === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (standardDeviation <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      const z = (x - mean) / standardDeviation;
      return numericResult(
        cumulative ? standardNormalCdf(z) : standardNormalPdf(z) / standardDeviation,
      );
    }
    if (fn === 'NORMSINV' || fn === 'NORM.S.INV') {
      const [probabilityOperand] = args;
      if (!probabilityOperand) return { kind: 'error', code: 15, text: '#VALUE!' };
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      if (probability === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (probability <= 0 || probability >= 1) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(inverseStandardNormal(probability));
    }
    if (fn === 'NORMINV' || fn === 'NORM.INV' || fn === 'LOGINV' || fn === 'LOGNORM.INV') {
      const [probabilityOperand, meanOperand, standardDeviationOperand] = args;
      if (!probabilityOperand || !meanOperand || !standardDeviationOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      if (probability === null || mean === null || standardDeviation === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (probability <= 0 || probability >= 1 || standardDeviation <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const value = mean + standardDeviation * inverseStandardNormal(probability);
      return numericResult(fn === 'LOGINV' || fn === 'LOGNORM.INV' ? Math.exp(value) : value);
    }
    if (fn === 'LOGNORMDIST' || fn === 'LOGNORM.DIST') {
      const [xOperand, meanOperand, standardDeviationOperand, cumulativeOperand] = args;
      if (!xOperand || !meanOperand || !standardDeviationOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const standardDeviation = readNumber(
        readOperand(standardDeviationOperand, rowOffset, colOffset),
      );
      const cumulative =
        fn === 'LOGNORMDIST'
          ? true
          : cumulativeOperand
            ? readLogical(readOperand(cumulativeOperand, rowOffset, colOffset))
            : null;
      if (x === null || mean === null || standardDeviation === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x <= 0 || standardDeviation <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      const z = (Math.log(x) - mean) / standardDeviation;
      return numericResult(
        cumulative ? standardNormalCdf(z) : standardNormalPdf(z) / (x * standardDeviation),
      );
    }
    if (fn === 'GAMMADIST' || fn === 'GAMMA.DIST') {
      const [xOperand, alphaOperand, betaOperand, cumulativeOperand] = args;
      if (!xOperand || !alphaOperand || !betaOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || alpha === null || beta === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x < 0 || alpha <= 0 || beta <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      if (!cumulative) {
        if (x === 0) {
          if (alpha === 1) return numericResult(1 / beta);
          return alpha > 1
            ? { kind: 'number', value: 0 }
            : { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult(
          Math.exp((alpha - 1) * Math.log(x) - x / beta - alpha * Math.log(beta) - logGamma(alpha)),
        );
      }
      const result = regularizedGammaP(alpha, x / beta);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'GAMMAINV' || fn === 'GAMMA.INV') {
      const [probabilityOperand, alphaOperand, betaOperand] = args;
      if (!probabilityOperand || !alphaOperand || !betaOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      if (probability === null || alpha === null || beta === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (probability <= 0 || probability >= 1 || alpha <= 0 || beta <= 0) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const result = inverseRegularizedGammaP(alpha, probability);
      return result === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : numericResult(result * beta);
    }
    if (fn === 'BETADIST' || fn === 'BETA.DIST') {
      const [xOperand, alphaOperand, betaOperand, cumulativeOperand, lowerOperand, upperOperand] =
        args;
      if (
        !xOperand ||
        !alphaOperand ||
        !betaOperand ||
        (fn === 'BETA.DIST' && !cumulativeOperand)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'BETADIST'
          ? true
          : readLogical(readOperand(cumulativeOperand as FormulaOperand, rowOffset, colOffset));
      const lower = lowerOperand ? readNumber(readOperand(lowerOperand, rowOffset, colOffset)) : 0;
      const upper = upperOperand ? readNumber(readOperand(upperOperand, rowOffset, colOffset)) : 1;
      if (
        x === null ||
        alpha === null ||
        beta === null ||
        cumulative === null ||
        lower === null ||
        upper === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (alpha <= 0 || beta <= 0 || lower >= upper || x < lower || x > upper) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const normalized = (x - lower) / (upper - lower);
      if (!cumulative) {
        if (normalized === 0 || normalized === 1) {
          const edgeDensity =
            normalized === 0 && alpha === 1
              ? Math.exp(logGamma(alpha + beta) - logGamma(alpha) - logGamma(beta)) /
                (upper - lower)
              : normalized === 1 && beta === 1
                ? Math.exp(logGamma(alpha + beta) - logGamma(alpha) - logGamma(beta)) /
                  (upper - lower)
                : null;
          return edgeDensity === null
            ? { kind: 'error', code: 6, text: '#NUM!' }
            : numericResult(edgeDensity);
        }
        return numericResult(
          Math.exp(
            (alpha - 1) * Math.log(normalized) +
              (beta - 1) * Math.log(1 - normalized) +
              logGamma(alpha + beta) -
              logGamma(alpha) -
              logGamma(beta),
          ) /
            (upper - lower),
        );
      }
      const result = regularizedBeta(normalized, alpha, beta);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'BETAINV' || fn === 'BETA.INV') {
      const [probabilityOperand, alphaOperand, betaOperand, lowerOperand, upperOperand] = args;
      if (!probabilityOperand || !alphaOperand || !betaOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const lower = lowerOperand ? readNumber(readOperand(lowerOperand, rowOffset, colOffset)) : 0;
      const upper = upperOperand ? readNumber(readOperand(upperOperand, rowOffset, colOffset)) : 1;
      if (
        probability === null ||
        alpha === null ||
        beta === null ||
        lower === null ||
        upper === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (probability <= 0 || probability >= 1 || alpha <= 0 || beta <= 0 || lower >= upper) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const result = inverseRegularizedBeta(probability, alpha, beta);
      return result === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : numericResult(lower + result * (upper - lower));
    }
    if (fn === 'FDIST' || fn === 'F.DIST' || fn === 'F.DIST.RT') {
      const [xOperand, degrees1Operand, degrees2Operand, cumulativeOperand] = args;
      if (
        !xOperand ||
        !degrees1Operand ||
        !degrees2Operand ||
        (fn === 'F.DIST' && !cumulativeOperand)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const degrees1 = readNumber(readOperand(degrees1Operand, rowOffset, colOffset));
      const degrees2 = readNumber(readOperand(degrees2Operand, rowOffset, colOffset));
      const cumulative =
        fn === 'F.DIST'
          ? readLogical(readOperand(cumulativeOperand as FormulaOperand, rowOffset, colOffset))
          : true;
      if (x === null || degrees1 === null || degrees2 === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const d1 = Math.trunc(degrees1);
      const d2 = Math.trunc(degrees2);
      if (x < 0 || d1 < 1 || d2 < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      const transformed = d1 * x === 0 ? 0 : (d1 * x) / (d1 * x + d2);
      if (!cumulative) {
        if (x === 0) {
          return d1 === 2
            ? { kind: 'number', value: 1 }
            : { kind: 'error', code: 6, text: '#NUM!' };
        }
        const halfD1 = d1 / 2;
        const halfD2 = d2 / 2;
        const logDensity =
          halfD1 * Math.log(d1 / d2) +
          (halfD1 - 1) * Math.log(x) -
          (halfD1 + halfD2) * Math.log(1 + (d1 * x) / d2) +
          logGamma(halfD1 + halfD2) -
          logGamma(halfD1) -
          logGamma(halfD2);
        return numericResult(Math.exp(logDensity));
      }
      const leftTail = regularizedBeta(transformed, d1 / 2, d2 / 2);
      if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(fn === 'F.DIST' ? leftTail : 1 - leftTail);
    }
    if (fn === 'FINV' || fn === 'F.INV' || fn === 'F.INV.RT') {
      const [probabilityOperand, degrees1Operand, degrees2Operand] = args;
      if (!probabilityOperand || !degrees1Operand || !degrees2Operand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const degrees1 = readNumber(readOperand(degrees1Operand, rowOffset, colOffset));
      const degrees2 = readNumber(readOperand(degrees2Operand, rowOffset, colOffset));
      if (probability === null || degrees1 === null || degrees2 === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const d1 = Math.trunc(degrees1);
      const d2 = Math.trunc(degrees2);
      if (probability <= 0 || probability >= 1 || d1 < 1 || d2 < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const leftTailProbability = fn === 'F.INV' ? probability : 1 - probability;
      const transformed = inverseRegularizedBeta(leftTailProbability, d1 / 2, d2 / 2);
      if (transformed === null || transformed >= 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      return numericResult((d2 * transformed) / (d1 * (1 - transformed)));
    }
    if (fn === 'TDIST' || fn === 'T.DIST' || fn === 'T.DIST.2T' || fn === 'T.DIST.RT') {
      const [xOperand, degreesOperand, cumulativeOrTailsOperand] = args;
      if (
        !xOperand ||
        !degreesOperand ||
        ((fn === 'TDIST' || fn === 'T.DIST') && !cumulativeOrTailsOperand)
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      if (x === null || degrees === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const d = Math.trunc(degrees);
      if (d < 1 || ((fn === 'TDIST' || fn === 'T.DIST.2T' || fn === 'T.DIST.RT') && x < 0)) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (fn === 'T.DIST') {
        const cumulative = readLogical(
          readOperand(cumulativeOrTailsOperand as FormulaOperand, rowOffset, colOffset),
        );
        if (cumulative === null) return { kind: 'error', code: 15, text: '#VALUE!' };
        if (!cumulative) return numericResult(studentTPdf(x, d));
        const leftTail = studentTCdf(x, d);
        return leftTail === null
          ? { kind: 'error', code: 6, text: '#NUM!' }
          : numericResult(leftTail);
      }
      const leftTail = studentTCdf(x, d);
      if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
      if (fn === 'T.DIST.RT') return numericResult(1 - leftTail);
      if (fn === 'T.DIST.2T') return numericResult(2 * (1 - leftTail));
      const tails = readNumber(
        readOperand(cumulativeOrTailsOperand as FormulaOperand, rowOffset, colOffset),
      );
      if (tails === null) return { kind: 'error', code: 15, text: '#VALUE!' };
      const tailCount = Math.trunc(tails);
      if (tailCount !== 1 && tailCount !== 2) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(tailCount === 1 ? 1 - leftTail : 2 * (1 - leftTail));
    }
    if (fn === 'TINV' || fn === 'T.INV' || fn === 'T.INV.2T') {
      const [probabilityOperand, degreesOperand] = args;
      if (!probabilityOperand || !degreesOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      if (probability === null || degrees === null)
        return { kind: 'error', code: 15, text: '#VALUE!' };
      const d = Math.trunc(degrees);
      if (probability <= 0 || probability >= 1 || d < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const leftTailProbability = fn === 'T.INV' ? probability : 1 - probability / 2;
      const result = inverseStudentTCdf(leftTailProbability, d);
      return result === null ? { kind: 'error', code: 6, text: '#NUM!' } : numericResult(result);
    }
    if (fn === 'CHIDIST' || fn === 'CHISQ.DIST' || fn === 'CHISQ.DIST.RT') {
      const [xOperand, degreesOperand, cumulativeOperand] = args;
      if (!xOperand || !degreesOperand || (fn === 'CHISQ.DIST' && !cumulativeOperand)) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'CHISQ.DIST'
          ? readLogical(readOperand(cumulativeOperand as FormulaOperand, rowOffset, colOffset))
          : true;
      if (x === null || degrees === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const df = Math.trunc(degrees);
      if (x < 0 || df < 1) return { kind: 'error', code: 6, text: '#NUM!' };
      const alpha = df / 2;
      if (!cumulative) {
        if (x === 0) {
          if (df === 2) return { kind: 'number', value: 0.5 };
          return df > 2 ? { kind: 'number', value: 0 } : { kind: 'error', code: 6, text: '#NUM!' };
        }
        return numericResult(Math.exp((alpha - 1) * Math.log(x / 2) - x / 2 - logGamma(alpha)) / 2);
      }
      const leftTail = regularizedGammaP(alpha, x / 2);
      if (leftTail === null) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(fn === 'CHISQ.DIST' ? leftTail : 1 - leftTail);
    }
    if (fn === 'CHIINV' || fn === 'CHISQ.INV' || fn === 'CHISQ.INV.RT') {
      const [probabilityOperand, degreesOperand] = args;
      if (!probabilityOperand || !degreesOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const degrees = readNumber(readOperand(degreesOperand, rowOffset, colOffset));
      if (probability === null || degrees === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const df = Math.trunc(degrees);
      if (probability <= 0 || probability >= 1 || df < 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      const leftTailProbability = fn === 'CHISQ.INV' ? probability : 1 - probability;
      const result = inverseRegularizedGammaP(df / 2, leftTailProbability);
      return result === null
        ? { kind: 'error', code: 6, text: '#NUM!' }
        : numericResult(result * 2);
    }
    if (fn === 'BINOMDIST' || fn === 'BINOM.DIST') {
      const [successesOperand, trialsOperand, probabilityOperand, cumulativeOperand] = args;
      if (!successesOperand || !trialsOperand || !probabilityOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const successes = readNumber(readOperand(successesOperand, rowOffset, colOffset));
      const trials = readNumber(readOperand(trialsOperand, rowOffset, colOffset));
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (successes === null || trials === null || probability === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const successCount = Math.trunc(successes);
      const trialCount = Math.trunc(trials);
      if (
        successCount < 0 ||
        trialCount < 0 ||
        successCount > trialCount ||
        probability < 0 ||
        probability > 1
      ) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (!cumulative) {
        return numericResult(binomialProbability(successCount, trialCount, probability));
      }
      let total = 0;
      for (let k = 0; k <= successCount; k += 1) {
        total += binomialProbability(k, trialCount, probability);
      }
      return numericResult(total);
    }
    if (fn === 'CRITBINOM' || fn === 'BINOM.INV') {
      const [trialsOperand, probabilityOperand, alphaOperand] = args;
      if (!trialsOperand || !probabilityOperand || !alphaOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const trials = readNumber(readOperand(trialsOperand, rowOffset, colOffset));
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      if (trials === null || probability === null || alpha === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const trialCount = Math.trunc(trials);
      if (trialCount < 0 || probability < 0 || probability > 1 || alpha <= 0 || alpha >= 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      let total = 0;
      for (let k = 0; k <= trialCount; k += 1) {
        total += binomialProbability(k, trialCount, probability);
        if (total >= alpha) return { kind: 'number', value: k };
      }
      return { kind: 'number', value: trialCount };
    }
    if (fn === 'NEGBINOMDIST' || fn === 'NEGBINOM.DIST') {
      const [failuresOperand, successesOperand, probabilityOperand, cumulativeOperand] = args;
      if (!failuresOperand || !successesOperand || !probabilityOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const failures = readNumber(readOperand(failuresOperand, rowOffset, colOffset));
      const successes = readNumber(readOperand(successesOperand, rowOffset, colOffset));
      const probability = readNumber(readOperand(probabilityOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'NEGBINOMDIST'
          ? false
          : cumulativeOperand
            ? readLogical(readOperand(cumulativeOperand, rowOffset, colOffset))
            : null;
      if (failures === null || successes === null || probability === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const failureCount = Math.trunc(failures);
      const successCount = Math.trunc(successes);
      if (failureCount < 0 || successCount < 1 || probability < 0 || probability > 1) {
        return { kind: 'error', code: 6, text: '#NUM!' };
      }
      if (!cumulative) {
        return numericResult(negativeBinomialProbability(failureCount, successCount, probability));
      }
      let total = 0;
      for (let k = 0; k <= failureCount; k += 1) {
        total += negativeBinomialProbability(k, successCount, probability);
      }
      return numericResult(total);
    }
    if (fn === 'HYPGEOMDIST' || fn === 'HYPGEOM.DIST') {
      const [
        sampleSuccessesOperand,
        sampleSizeOperand,
        populationSuccessesOperand,
        populationSizeOperand,
        cumulativeOperand,
      ] = args;
      if (
        !sampleSuccessesOperand ||
        !sampleSizeOperand ||
        !populationSuccessesOperand ||
        !populationSizeOperand
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const sampleSuccesses = readNumber(readOperand(sampleSuccessesOperand, rowOffset, colOffset));
      const sampleSize = readNumber(readOperand(sampleSizeOperand, rowOffset, colOffset));
      const populationSuccesses = readNumber(
        readOperand(populationSuccessesOperand, rowOffset, colOffset),
      );
      const populationSize = readNumber(readOperand(populationSizeOperand, rowOffset, colOffset));
      const cumulative =
        fn === 'HYPGEOMDIST'
          ? false
          : cumulativeOperand
            ? readLogical(readOperand(cumulativeOperand, rowOffset, colOffset))
            : null;
      if (
        sampleSuccesses === null ||
        sampleSize === null ||
        populationSuccesses === null ||
        populationSize === null ||
        cumulative === null
      ) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = Math.trunc(sampleSuccesses);
      const n = Math.trunc(sampleSize);
      const m = Math.trunc(populationSuccesses);
      const bigN = Math.trunc(populationSize);
      const valid =
        x >= 0 && n >= 0 && m >= 0 && bigN >= 0 && x <= n && x <= m && n <= bigN && m <= bigN;
      if (!valid || n - x > bigN - m) return { kind: 'error', code: 6, text: '#NUM!' };
      if (!cumulative) return numericResult(hypergeometricProbability(x, n, m, bigN));
      let total = 0;
      const minSuccess = Math.max(0, n - (bigN - m));
      for (let k = minSuccess; k <= x; k += 1) {
        total += hypergeometricProbability(k, n, m, bigN);
      }
      return numericResult(total);
    }
    if (fn === 'POISSON' || fn === 'POISSON.DIST') {
      const [xOperand, meanOperand, cumulativeOperand] = args;
      if (!xOperand || !meanOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const mean = readNumber(readOperand(meanOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || mean === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const count = Math.trunc(x);
      if (count < 0 || mean < 0) return { kind: 'error', code: 6, text: '#NUM!' };
      if (!cumulative) return numericResult(poissonProbability(count, mean));
      let total = 0;
      for (let k = 0; k <= count; k += 1) total += poissonProbability(k, mean);
      return numericResult(total);
    }
    if (fn === 'EXPONDIST' || fn === 'EXPON.DIST') {
      const [xOperand, lambdaOperand, cumulativeOperand] = args;
      if (!xOperand || !lambdaOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const lambda = readNumber(readOperand(lambdaOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || lambda === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x < 0 || lambda <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      return numericResult(cumulative ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x));
    }
    if (fn === 'WEIBULL' || fn === 'WEIBULL.DIST') {
      const [xOperand, alphaOperand, betaOperand, cumulativeOperand] = args;
      if (!xOperand || !alphaOperand || !betaOperand || !cumulativeOperand) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      const x = readNumber(readOperand(xOperand, rowOffset, colOffset));
      const alpha = readNumber(readOperand(alphaOperand, rowOffset, colOffset));
      const beta = readNumber(readOperand(betaOperand, rowOffset, colOffset));
      const cumulative = readLogical(readOperand(cumulativeOperand, rowOffset, colOffset));
      if (x === null || alpha === null || beta === null || cumulative === null) {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
      if (x < 0 || alpha <= 0 || beta <= 0) return { kind: 'error', code: 6, text: '#NUM!' };
      const scaled = (x / beta) ** alpha;
      return numericResult(
        cumulative
          ? 1 - Math.exp(-scaled)
          : (alpha / beta) * (x / beta) ** (alpha - 1) * Math.exp(-scaled),
      );
    }
    return undefined;
  };
  return { probabilityFunction };
}
