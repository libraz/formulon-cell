const erf = (value: number): number => {
  const sign = value < 0 ? -1 : 1;
  const x = Math.abs(value);
  const t = 1 / (1 + 0.5 * x);
  let polynomial = 0.17087277;
  polynomial = -0.82215223 + t * polynomial;
  polynomial = 1.48851587 + t * polynomial;
  polynomial = -1.13520398 + t * polynomial;
  polynomial = 0.27886807 + t * polynomial;
  polynomial = -0.18628806 + t * polynomial;
  polynomial = 0.09678418 + t * polynomial;
  polynomial = 0.37409196 + t * polynomial;
  polynomial = 1.00002368 + t * polynomial;
  const tau = t * Math.exp(-x * x - 1.26551223 + t * polynomial);
  return sign * (1 - tau);
};
const standardNormalCdf = (z: number): number => 0.5 * (1 + erf(z / Math.SQRT2));
const standardNormalPdf = (z: number): number => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
const inverseStandardNormal = (probability: number): number => {
  const pick = (values: number[], index: number): number => values[index] ?? 0;
  const a = [
    -39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716,
    2.506628277459239,
  ];
  const b = [
    -54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972,
    -13.28068155288572,
  ];
  const c = [
    -0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734,
    4.374664141464968, 2.938163982698783,
  ];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const low = 0.02425;
  const high = 1 - low;
  if (probability < low) {
    const q = Math.sqrt(-2 * Math.log(probability));
    const numerator =
      ((((pick(c, 0) * q + pick(c, 1)) * q + pick(c, 2)) * q + pick(c, 3)) * q + pick(c, 4)) * q +
      pick(c, 5);
    const denominator = (((pick(d, 0) * q + pick(d, 1)) * q + pick(d, 2)) * q + pick(d, 3)) * q + 1;
    return numerator / denominator;
  }
  if (probability > high) {
    const q = Math.sqrt(-2 * Math.log(1 - probability));
    const numerator =
      ((((pick(c, 0) * q + pick(c, 1)) * q + pick(c, 2)) * q + pick(c, 3)) * q + pick(c, 4)) * q +
      pick(c, 5);
    const denominator = (((pick(d, 0) * q + pick(d, 1)) * q + pick(d, 2)) * q + pick(d, 3)) * q + 1;
    return -(numerator / denominator);
  }
  const q = probability - 0.5;
  const r = q * q;
  const numerator =
    (((((pick(a, 0) * r + pick(a, 1)) * r + pick(a, 2)) * r + pick(a, 3)) * r + pick(a, 4)) * r +
      pick(a, 5)) *
    q;
  const denominator =
    ((((pick(b, 0) * r + pick(b, 1)) * r + pick(b, 2)) * r + pick(b, 3)) * r + pick(b, 4)) * r + 1;
  return numerator / denominator;
};
const binomialProbability = (successes: number, trials: number, probability: number): number => {
  if (probability === 0) return successes === 0 ? 1 : 0;
  if (probability === 1) return successes === trials ? 1 : 0;
  let coefficient = 1;
  const choose = Math.min(successes, trials - successes);
  for (let i = 1; i <= choose; i += 1) {
    coefficient *= (trials - choose + i) / i;
  }
  return coefficient * probability ** successes * (1 - probability) ** (trials - successes);
};
const poissonProbability = (x: number, mean: number): number => {
  if (mean === 0) return x === 0 ? 1 : 0;
  let factorial = 1;
  for (let i = 2; i <= x; i += 1) factorial *= i;
  return (Math.exp(-mean) * mean ** x) / factorial;
};
const factorial = (value: number): number => {
  let result = 1;
  for (let i = 2; i <= value; i += 1) result *= i;
  return result;
};
const doubleFactorial = (value: number): number => {
  let result = 1;
  for (let i = value; i > 1; i -= 2) result *= i;
  return result;
};
const logGamma = (value: number): number => {
  const coefficients = [
    676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406,
    12.507343278686905, -0.13857109526572012, 0.000009984369578019572, 0.00000015056327351493116,
  ];
  if (value < 0.5) {
    return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * value)) - logGamma(1 - value);
  }
  const z = value - 1;
  let x = 0.9999999999998099;
  for (let i = 0; i < coefficients.length; i += 1) {
    x += (coefficients[i] as number) / (z + i + 1);
  }
  const t = z + coefficients.length - 0.5;
  return Math.log(Math.sqrt(2 * Math.PI)) + (z + 0.5) * Math.log(t) - t + Math.log(x);
};
const gamma = (value: number): number | null => {
  if (value === 0 || (value < 0 && Number.isInteger(value))) return null;
  if (value < 0.5) {
    return Math.PI / (Math.sin(Math.PI * value) * Math.exp(logGamma(1 - value)));
  }
  return Math.exp(logGamma(value));
};
const regularizedGammaP = (alpha: number, x: number): number | null => {
  if (x <= 0) return 0;
  const epsilon = 1e-12;
  const maxIterations = 100;
  const tiny = 1e-300;
  const logTerm = alpha * Math.log(x) - x - logGamma(alpha);
  if (x < alpha + 1) {
    let sum = 1 / alpha;
    let term = sum;
    for (let n = 1; n <= maxIterations; n += 1) {
      term *= x / (alpha + n);
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * epsilon) {
        return Math.exp(logTerm) * sum;
      }
    }
    return null;
  }
  let b = x + 1 - alpha;
  let c = 1 / tiny;
  let d = 1 / Math.max(b, tiny);
  let h = d;
  for (let i = 1; i <= maxIterations; i += 1) {
    const an = -i * (i - alpha);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < tiny) d = tiny;
    c = b + an / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < epsilon) {
      return 1 - Math.exp(logTerm) * h;
    }
  }
  return null;
};
const inverseRegularizedGammaP = (alpha: number, probability: number): number | null => {
  let low = 0;
  let high = Math.max(1, alpha);
  for (let i = 0; i < 100; i += 1) {
    const value = regularizedGammaP(alpha, high);
    if (value === null) return null;
    if (value >= probability) break;
    high *= 2;
    if (!Number.isFinite(high)) return null;
  }
  for (let i = 0; i < 100; i += 1) {
    const mid = (low + high) / 2;
    const value = regularizedGammaP(alpha, mid);
    if (value === null) return null;
    if (value < probability) low = mid;
    else high = mid;
  }
  return (low + high) / 2;
};
const betaContinuedFraction = (x: number, alpha: number, beta: number): number | null => {
  const maxIterations = 100;
  const epsilon = 3e-14;
  const tiny = 1e-300;
  const qab = alpha + beta;
  const qap = alpha + 1;
  const qam = alpha - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < tiny) d = tiny;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= maxIterations; m += 1) {
    const m2 = 2 * m;
    let aa = (m * (beta - m) * x) / ((qam + m2) * (alpha + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    h *= d * c;
    aa = (-(alpha + m) * (qab + m) * x) / ((alpha + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < epsilon) return h;
  }
  return null;
};
const regularizedBeta = (x: number, alpha: number, beta: number): number | null => {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const logBt =
    logGamma(alpha + beta) -
    logGamma(alpha) -
    logGamma(beta) +
    alpha * Math.log(x) +
    beta * Math.log(1 - x);
  if (x < (alpha + 1) / (alpha + beta + 2)) {
    const fraction = betaContinuedFraction(x, alpha, beta);
    return fraction === null ? null : (Math.exp(logBt) * fraction) / alpha;
  }
  const fraction = betaContinuedFraction(1 - x, beta, alpha);
  return fraction === null ? null : 1 - (Math.exp(logBt) * fraction) / beta;
};
const inverseRegularizedBeta = (
  probability: number,
  alpha: number,
  beta: number,
): number | null => {
  let low = 0;
  let high = 1;
  for (let i = 0; i < 100; i += 1) {
    const mid = (low + high) / 2;
    const value = regularizedBeta(mid, alpha, beta);
    if (value === null) return null;
    if (value < probability) low = mid;
    else high = mid;
  }
  return (low + high) / 2;
};
const studentTCdf = (x: number, degrees: number): number | null => {
  if (x === 0) return 0.5;
  const betaInput = degrees / (degrees + x * x);
  const betaValue = regularizedBeta(betaInput, degrees / 2, 0.5);
  if (betaValue === null) return null;
  return x > 0 ? 1 - betaValue / 2 : betaValue / 2;
};
const studentTPdf = (x: number, degrees: number): number =>
  Math.exp(
    logGamma((degrees + 1) / 2) -
      logGamma(degrees / 2) -
      0.5 * Math.log(degrees * Math.PI) -
      ((degrees + 1) / 2) * Math.log(1 + (x * x) / degrees),
  );
const inverseStudentTCdf = (probability: number, degrees: number): number | null => {
  let low = -1;
  let high = 1;
  for (let i = 0; i < 100; i += 1) {
    const lowValue = studentTCdf(low, degrees);
    const highValue = studentTCdf(high, degrees);
    if (lowValue === null || highValue === null) return null;
    if (lowValue <= probability && highValue >= probability) break;
    low *= 2;
    high *= 2;
    if (!Number.isFinite(low) || !Number.isFinite(high)) return null;
  }
  for (let i = 0; i < 100; i += 1) {
    const mid = (low + high) / 2;
    const value = studentTCdf(mid, degrees);
    if (value === null) return null;
    if (value < probability) low = mid;
    else high = mid;
  }
  return (low + high) / 2;
};
const combination = (n: number, k: number): number => {
  const choose = Math.min(k, n - k);
  let result = 1;
  for (let i = 1; i <= choose; i += 1) result *= (n - choose + i) / i;
  return result;
};
const negativeBinomialProbability = (
  failures: number,
  successes: number,
  probability: number,
): number =>
  combination(failures + successes - 1, failures) *
  probability ** successes *
  (1 - probability) ** failures;
const hypergeometricProbability = (
  sampleSuccesses: number,
  sampleSize: number,
  populationSuccesses: number,
  populationSize: number,
): number =>
  (combination(populationSuccesses, sampleSuccesses) *
    combination(populationSize - populationSuccesses, sampleSize - sampleSuccesses)) /
  combination(populationSize, sampleSize);

export {
  binomialProbability,
  combination,
  doubleFactorial,
  erf,
  factorial,
  gamma,
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
};
