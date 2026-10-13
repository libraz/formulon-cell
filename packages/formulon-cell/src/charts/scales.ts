export interface NumericDomain {
  readonly min: number;
  readonly max: number;
}

const finiteOr = (value: number, fallback: number): number =>
  Number.isFinite(value) ? value : fallback;

const cleanZero = (value: number): number => (value === 0 ? 0 : value);

const safePadding = (value: number): number => {
  const absolute = Math.abs(value);
  if (absolute === 0) return 1;
  const tenth = absolute * 0.1;
  if (Number.isFinite(tenth) && tenth > 0) return tenth;
  // Multiplication can underflow for subnormal values.  A single minimum
  // subnormal is still enough to separate the two sides of the domain,
  // whereas falling back to a value near MAX_VALUE makes tiny data appear
  // indistinguishable from zero.
  return Number.MIN_VALUE;
};

/** Return a stable domain even when the source is empty, constant, or invalid. */
export function numericDomain(
  values: readonly number[],
  fallback: NumericDomain = { min: 0, max: 1 },
): NumericDomain {
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (const value of values) {
    if (!Number.isFinite(value)) continue;
    min = Math.min(min, value);
    max = Math.max(max, value);
  }
  if (!Number.isFinite(min) || !Number.isFinite(max)) return fallback;
  if (min === max) {
    const padding = safePadding(min);
    if (min > 0) min -= padding;
    else if (min < 0) max += padding;
    else {
      min = -padding;
      max = padding;
    }
  }
  return { min, max };
}

/** Compute a compact 1/2/5-style axis interval. */
export function niceStep(min: number, max: number, targetCount = 5): number {
  const lo = Math.min(finiteOr(min, 0), finiteOr(max, 1));
  const hi = Math.max(finiteOr(min, 0), finiteOr(max, 1));
  const directSpan = hi - lo;
  const span = Number.isFinite(directSpan)
    ? directSpan || 1
    : Math.max(Math.abs(lo), Math.abs(hi), 1);
  const target = Math.max(2, Math.floor(finiteOr(targetCount, 5)));
  const raw = Math.max(Number.MIN_VALUE, span / target);
  const exponent = Math.floor(Math.log10(raw));
  const powerCandidate = Number.isFinite(exponent) ? 10 ** exponent : Number.MAX_VALUE;
  const power =
    powerCandidate === 0
      ? Number.MIN_VALUE
      : Number.isFinite(powerCandidate)
        ? powerCandidate
        : Number.MAX_VALUE;
  const fraction = raw / power;
  const multiplier = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
  const step = multiplier * power;
  return Number.isFinite(step) && step > 0 ? step : power;
}

/** Expand a domain to human-friendly tick boundaries. */
export function niceDomain(min: number, max: number, targetCount = 5): NumericDomain {
  let lo = finiteOr(min, 0);
  let hi = finiteOr(max, 1);
  if (lo > hi) [lo, hi] = [hi, lo];
  if (lo === hi) {
    const padding = safePadding(lo);
    if (lo > 0) lo -= padding;
    else if (lo < 0) hi += padding;
    else {
      lo = -padding;
      hi = padding;
    }
  }
  const step = niceStep(lo, hi, targetCount);
  const lowerIndex = Math.floor(lo / step);
  const upperIndex = Math.ceil(hi / step);
  const lower = lowerIndex * step;
  const upper = upperIndex * step;
  return {
    min: cleanZero(Number.isFinite(lower) ? lower : lo),
    max: cleanZero(Number.isFinite(upper) ? upper : hi),
  };
}

/** Generate deterministic nice ticks constrained to the supplied domain. */
export function niceTicks(min: number, max: number, targetCount = 5): readonly number[] {
  let lo = finiteOr(min, 0);
  let hi = finiteOr(max, 1);
  if (lo > hi) [lo, hi] = [hi, lo];
  if (lo === hi) return [cleanZero(lo)];
  const step = niceStep(lo, hi, targetCount);
  const first = Math.ceil(lo / step);
  const last = Math.floor(hi / step);
  const ticks: number[] = [];
  if (Number.isFinite(first) && Number.isFinite(last)) {
    for (let index = first; index <= last && ticks.length < 100; index += 1) {
      const tick = index * step;
      if (!Number.isFinite(tick)) break;
      if (tick >= lo && tick <= hi) ticks.push(cleanZero(Number(tick.toPrecision(12))));
    }
  }
  if (ticks.length === 0) {
    ticks.push(cleanZero(lo));
    if (hi !== lo) ticks.push(cleanZero(hi));
  }
  return ticks;
}

/**
 * Generate major ticks from the domain minimum at an explicit positive unit.
 * At most 101 values are returned; dense intervals are thinned by an integer
 * multiple of the requested unit so every retained tick remains deterministic.
 */
export function majorTicks(
  min: number,
  max: number,
  majorUnit: number | undefined,
): readonly number[] {
  let lo = finiteOr(min, 0);
  let hi = finiteOr(max, 1);
  if (lo > hi) [lo, hi] = [hi, lo];
  if (majorUnit === undefined || !Number.isFinite(majorUnit) || majorUnit <= 0 || lo === hi) {
    return niceTicks(lo, hi, 5);
  }
  const span = hi - lo;
  if (!Number.isFinite(span) || span <= 0) return niceTicks(lo, hi, 5);
  const ratio = span / majorUnit;
  if (!Number.isFinite(ratio)) return niceTicks(lo, hi, 5);
  // Decimal units such as 0.1 are commonly represented just below their
  // mathematical value (0.3 / 0.1 === 2.9999999999999996).  Allow one small
  // floating-point ulp when counting the final tick, while keeping the
  // resulting values clipped to the requested domain below.
  const ratioTolerance = Math.max(
    Number.MIN_VALUE,
    Number.EPSILON * 16 * Math.max(1, Math.abs(ratio)),
  );
  const estimatedCount = Math.floor(ratio + ratioTolerance) + 1;
  const stride = estimatedCount > 101 ? Math.ceil(estimatedCount / 101) : 1;
  const step = majorUnit * stride;
  if (!Number.isFinite(step) || step <= 0 || lo + step === lo) {
    return [cleanZero(lo)];
  }
  const stepRatio = span / step;
  const stepRatioTolerance = Math.max(
    Number.MIN_VALUE,
    Number.EPSILON * 16 * Math.max(1, Math.abs(stepRatio)),
  );
  const count = Math.min(100, Math.floor(stepRatio + stepRatioTolerance));
  const ticks: number[] = [];
  for (let index = 0; index <= count; index += 1) {
    let tick = lo + step * index;
    if (!Number.isFinite(tick)) break;
    const tickTolerance = Math.max(
      Number.MIN_VALUE,
      Number.EPSILON * 16 * Math.max(1, Math.abs(lo), Math.abs(hi), Math.abs(tick)),
    );
    if (tick < lo - tickTolerance || tick > hi + tickTolerance) break;
    if (tick < lo) tick = lo;
    if (tick > hi) tick = hi;
    if (index > 0 && tick === ticks[index - 1]) break;
    ticks.push(cleanZero(tick));
  }
  return ticks.length > 0 ? ticks : [cleanZero(lo)];
}

/** Map a finite value into a screen range without ever returning NaN/Infinity. */
export function linearScale(
  value: number | null,
  domainMin: number,
  domainMax: number,
  rangeMin: number,
  rangeMax: number,
): number | null {
  if (value === null || !Number.isFinite(value)) return null;
  const lo = finiteOr(domainMin, 0);
  const hi = finiteOr(domainMax, 1);
  let fraction = 0.5;
  if (lo !== hi) {
    const span = hi - lo;
    const offset = value - lo;
    if (Number.isFinite(span) && Number.isFinite(offset) && Math.abs(span) >= Number.MIN_VALUE) {
      fraction = offset / span;
    } else {
      const magnitude = Math.max(Math.abs(lo), Math.abs(hi), Math.abs(value), Number.MIN_VALUE);
      const normalizedValue = value / magnitude;
      const normalizedLo = lo / magnitude;
      const normalizedHi = hi / magnitude;
      const normalizedSpan = normalizedHi - normalizedLo;
      fraction = normalizedSpan === 0 ? 0.5 : (normalizedValue - normalizedLo) / normalizedSpan;
    }
  }
  const rangeLo = finiteOr(rangeMin, 0);
  const rangeHi = finiteOr(rangeMax, 1);
  const rangeSpan = rangeHi - rangeLo;
  const mapped = Number.isFinite(rangeSpan)
    ? rangeLo + fraction * rangeSpan
    : rangeLo * (1 - fraction) + rangeHi * fraction;
  if (Number.isFinite(mapped)) return mapped;
  if (fraction <= 0) return rangeLo;
  if (fraction >= 1) return rangeHi;
  return (rangeLo + rangeHi) / 2;
}

/** Format a tick label without locale or browser dependencies. */
export function formatTick(value: number, percent = false): string {
  if (!Number.isFinite(value)) return '';
  if (percent) return `${Math.round(value * 100)}%`;
  const absolute = Math.abs(value);
  if (absolute !== 0 && (absolute >= 1e6 || absolute < 1e-4)) {
    return value.toExponential(2).replace('e+', 'e');
  }
  return Number(value.toPrecision(6)).toString();
}
