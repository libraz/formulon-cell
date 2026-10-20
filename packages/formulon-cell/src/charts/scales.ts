export interface NumericDomain {
  readonly min: number;
  readonly max: number;
}

/** Return a usable logarithmic base, or leave the axis on its linear path. */
export const validateLogBase = (value: number | undefined): number | undefined =>
  Number.isFinite(value) && (value as number) >= 2 && (value as number) <= 1000
    ? (value as number)
    : undefined;

const positivePower = (base: number, exponent: number): number => {
  if (!Number.isFinite(base) || base <= 1 || !Number.isFinite(exponent)) return 1;
  const value = base ** exponent;
  if (Number.isFinite(value) && value > 0) return value;
  if (value === 0 || value < 0) return Number.MIN_VALUE;
  return Number.MAX_VALUE;
};

const powerExponent = (value: number, base: number): number => {
  const raw = Math.log(value) / Math.log(base);
  return Number.isFinite(raw) ? raw : 0;
};

const powerBoundary = (value: number, base: number, lower: boolean): number => {
  const exponent = powerExponent(value, base);
  const nearestExponent = Math.round(exponent);
  if (positivePower(base, nearestExponent) === value) {
    return positivePower(base, nearestExponent);
  }
  let candidate = positivePower(base, lower ? Math.floor(exponent) : Math.ceil(exponent));
  if (lower && candidate > value) {
    candidate = positivePower(base, Math.floor(exponent) - 1);
  } else if (!lower && candidate < value) {
    candidate = positivePower(base, Math.ceil(exponent) + 1);
  }
  return candidate;
};

const strictPowerBoundary = (value: number, base: number, lower: boolean): number => {
  const exponent = powerExponent(value, base);
  const nearestExponent = Math.round(exponent);
  if (positivePower(base, nearestExponent) === value) {
    return positivePower(base, nearestExponent + (lower ? -1 : 1));
  }
  const candidate = powerBoundary(value, base, lower);
  if ((lower && candidate < value) || (!lower && candidate > value)) return candidate;
  return positivePower(base, Math.round(exponent) + (lower ? -1 : 1));
};

const widenedPositiveDomain = (value: number, base: number): NumericDomain => {
  const exponent = powerExponent(value, base);
  const nearestExponent = Math.round(exponent);
  const exactPower = positivePower(base, nearestExponent) === value;
  let min = exactPower
    ? positivePower(base, nearestExponent - 1)
    : powerBoundary(value, base, true);
  let max = exactPower
    ? positivePower(base, nearestExponent + 1)
    : powerBoundary(value, base, false);
  if (!(min < value)) min = value === Number.MIN_VALUE ? value : value / base;
  if (!(max > value)) max = value === Number.MAX_VALUE ? value : value * base;
  if (!(min > 0) || !Number.isFinite(min)) min = Number.MIN_VALUE;
  if (!(max > 0) || !Number.isFinite(max)) max = Number.MAX_VALUE;
  if (min === max) {
    if (value < Number.MAX_VALUE) max = Math.min(Number.MAX_VALUE, value * (1 + Number.EPSILON));
    if (min === max && value > Number.MIN_VALUE) min = value * (1 - Number.EPSILON);
  }
  return { min, max };
};

/**
 * Resolve a positive domain for a logarithmic axis.
 *
 * Non-positive and non-finite values/bounds are ignored. Automatic bounds are
 * expanded to surrounding powers; explicit positive bounds remain exact.
 */
export function resolvePositivePowerDomain(
  values: readonly number[],
  explicitMin: number | undefined,
  explicitMax: number | undefined,
  baseInput: number,
): NumericDomain {
  const base = validateLogBase(baseInput) ?? 10;
  const positiveValues = values.filter((value) => Number.isFinite(value) && value > 0);
  const hasExplicitMin = Number.isFinite(explicitMin) && (explicitMin as number) > 0;
  const hasExplicitMax = Number.isFinite(explicitMax) && (explicitMax as number) > 0;

  if (positiveValues.length === 0 && !hasExplicitMin && !hasExplicitMax) {
    return { min: 1, max: base };
  }

  const observedMin = positiveValues.reduce(
    (current, value) => Math.min(current, value),
    Number.POSITIVE_INFINITY,
  );
  const observedMax = positiveValues.reduce(
    (current, value) => Math.max(current, value),
    Number.NEGATIVE_INFINITY,
  );
  let min = hasExplicitMin ? (explicitMin as number) : observedMin;
  let max = hasExplicitMax ? (explicitMax as number) : observedMax;

  if (hasExplicitMin && hasExplicitMax) {
    if (min > max) [min, max] = [max, min];
    if (min === max) return widenedPositiveDomain(min, base);
    return { min, max };
  }

  if (!(min > 0) || !Number.isFinite(min)) {
    min = hasExplicitMax ? (explicitMax as number) : 1;
  }
  if (!(max > 0) || !Number.isFinite(max)) {
    max = hasExplicitMin ? (explicitMin as number) : base;
  }

  if (hasExplicitMin && !hasExplicitMax) {
    if (!(max > min)) max = strictPowerBoundary(min, base, false);
    else max = powerBoundary(max, base, false);
  } else if (!hasExplicitMin && hasExplicitMax) {
    if (!(min < max)) min = strictPowerBoundary(max, base, true);
    else min = powerBoundary(min, base, true);
  } else if (min === max) {
    return widenedPositiveDomain(min, base);
  } else {
    min = powerBoundary(min, base, true);
    max = powerBoundary(max, base, false);
  }

  if (!(min > 0) || !Number.isFinite(min)) min = Number.MIN_VALUE;
  if (!(max > 0) || !Number.isFinite(max)) max = Number.MAX_VALUE;
  if (min === max) {
    if (hasExplicitMin && !hasExplicitMax) return widenedPositiveDomain(min, base);
    if (!hasExplicitMin && hasExplicitMax) return widenedPositiveDomain(min, base);
    return widenedPositiveDomain(min, base);
  }
  if (min > max) [min, max] = [max, min];
  return { min, max };
}

/** Generate ascending base-power ticks, bounded to 101 finite values. */
export function powerTicks(
  minInput: number,
  maxInput: number,
  baseInput: number,
): readonly number[] {
  const base = validateLogBase(baseInput) ?? 10;
  let min = Number.isFinite(minInput) && minInput > 0 ? minInput : 1;
  let max = Number.isFinite(maxInput) && maxInput > 0 ? maxInput : base;
  if (min > max) [min, max] = [max, min];
  if (min === max) {
    const exponent = Math.round(powerExponent(min, base));
    return positivePower(base, exponent) === min ? [min] : [];
  }

  const lowerLog = Math.log(min) / Math.log(base);
  const upperLog = Math.log(max) / Math.log(base);
  if (!Number.isFinite(lowerLog) || !Number.isFinite(upperLog)) return [];
  const tolerance = Number.EPSILON * 32 * Math.max(1, Math.abs(lowerLog), Math.abs(upperLog));
  const first = Math.ceil(lowerLog - tolerance);
  const last = Math.floor(upperLog + tolerance);
  if (!Number.isFinite(first) || !Number.isFinite(last) || last < first) return [];

  const count = last - first + 1;
  const stride = count > 101 ? Math.ceil(count / 101) : 1;
  const ticks: number[] = [];
  for (let exponent = first; exponent <= last; exponent += stride) {
    const tick = positivePower(base, exponent);
    if (!(tick > 0) || !Number.isFinite(tick)) continue;
    if (tick < min || tick > max) continue;
    if (ticks.length === 0 || tick > (ticks[ticks.length - 1] as number)) ticks.push(tick);
    if (ticks.length >= 101) break;
  }
  return ticks;
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

/**
 * The major unit Excel picks for an automatic axis over `span`: the largest 1,
 * 2 or 5 times a power of ten that still divides the span into five or more
 * intervals (0 to 16 by 2, 0 to 120 by 20).
 */
export function automaticMajorUnit(span: number): number {
  if (!Number.isFinite(span) || span <= 0) return 1;
  const power = 10 ** Math.floor(Math.log10(span / 5));
  if (!Number.isFinite(power) || power <= 0) return span;
  for (const multiplier of [10, 5, 2, 1]) {
    const unit = multiplier * power;
    if (span / unit >= 5 - 1e-9) return unit;
  }
  return power;
}

/**
 * Excel's automatic bounds of a value axis over data in [`min`, `max`]. A
 * positive axis starts at zero unless the data spans less than a sixth of its
 * maximum, in which case it starts half the span below the minimum; a free
 * side gets a twentieth of the span of headroom and is rounded out to the
 * major unit. A side given in `fixed` keeps that value.
 */
export function automaticAxisDomain(
  min: number,
  max: number,
  fixed: { readonly min?: number; readonly max?: number } = {},
): NumericDomain & { readonly majorUnit: number } {
  let lo = finiteOr(Math.min(min, max), 0);
  let hi = finiteOr(Math.max(min, max), 1);
  let padLo = true;
  let padHi = true;
  if (lo === 0 && hi === 0 && fixed.max === undefined) {
    hi = 1;
  } else if (fixed.min === undefined && lo >= 0 && hi > 0) {
    padLo = false;
    lo = lo === hi || (hi - lo) / hi > 1 / 6 ? 0 : lo - (hi - lo) / 2;
  } else if (fixed.max === undefined && hi <= 0 && lo < 0) {
    padHi = false;
    hi = lo === hi || (hi - lo) / -lo > 1 / 6 ? 0 : hi + (hi - lo) / 2;
  }
  if (fixed.min !== undefined) lo = fixed.min;
  if (fixed.max !== undefined) hi = fixed.max;
  const span = hi - lo;
  if (!Number.isFinite(span) || span <= 0) {
    const nice = niceDomain(lo, hi, 5);
    return { ...nice, majorUnit: niceStep(nice.min, nice.max, 5) };
  }
  // Headroom and rounding that would overflow leave the side at the data.
  const within = (value: number, fallback: number): number =>
    Number.isFinite(value) ? value : fallback;
  const bottom = within(fixed.min === undefined && padLo && lo < 0 ? lo - span / 20 : lo, lo);
  const top = within(fixed.max === undefined && padHi && hi > 0 ? hi + span / 20 : hi, hi);
  const majorUnit = automaticMajorUnit(within(top - bottom, span));
  const round = (value: number): number => cleanZero(Number(value.toPrecision(12)));
  return {
    min:
      fixed.min === undefined ? within(round(Math.floor(bottom / majorUnit) * majorUnit), lo) : lo,
    max:
      fixed.max === undefined
        ? within(round(Math.ceil(top / majorUnit - 1e-9) * majorUnit), hi)
        : hi,
    majorUnit,
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

/**
 * Project a positive value in log space without clamping the resulting
 * fraction. Endpoint checks keep adjacent extreme values from collapsing to
 * a midpoint when their natural logarithms round to the same number.
 */
export function unclampedLogProjection(
  value: number,
  domainMin: number,
  domainMax: number,
  rangeMin: number,
  rangeMax: number,
  reverse = false,
): number {
  const from = reverse ? rangeMax : rangeMin;
  const to = reverse ? rangeMin : rangeMax;
  const lo = domainMin;
  const hi = domainMax;
  if (
    !Number.isFinite(value) ||
    !(value > 0) ||
    !Number.isFinite(lo) ||
    !(lo > 0) ||
    !Number.isFinite(hi) ||
    !(hi > 0) ||
    lo === hi
  ) {
    return linearScale(value, lo, hi, from, to) ?? from;
  }
  if (value === lo) return from;
  if (value === hi) return to;
  const logMin = Math.log(lo);
  const logMax = Math.log(hi);
  const logSpan = logMax - logMin;
  if (!(logSpan > 0) || !Number.isFinite(logSpan)) {
    return linearScale(value, lo, hi, from, to) ?? from;
  }
  const logValue = Math.log(value);
  const fraction = (logValue - logMin) / logSpan;
  if (!Number.isFinite(fraction)) return linearScale(value, lo, hi, from, to) ?? from;
  const mapped = from + fraction * (to - from);
  if (Number.isFinite(mapped)) return mapped;
  if (fraction <= 0) return from;
  if (fraction >= 1) return to;
  return (from + to) / 2;
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
