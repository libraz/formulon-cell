import type { ChartPlotBounds } from './types.js';

/** A finite data point, or a source-order break/blank. */
export type ScatterLineSample =
  | {
      readonly kind: 'point';
      readonly x: number;
      readonly y: number;
      readonly sourceIndex: number;
    }
  | {
      readonly kind: 'blank';
      readonly sourceIndex: number;
    }
  | {
      readonly kind: 'break';
      readonly sourceIndex: number;
    };

export interface ScatterLineDomain {
  readonly min: number;
  readonly max: number;
}

export interface ScatterLineDomains {
  readonly x: ScatterLineDomain;
  readonly y: ScatterLineDomain;
}

/** Optional screen projection flags applied after raw-domain clipping. */
export interface ScatterLineProjection {
  readonly reverseX?: boolean;
  readonly reverseY?: boolean;
}

/** A clipped source-order segment ready to become one SVG line primitive. */
export interface ScatterLineSegment {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
  readonly startIndex: number;
  readonly endIndex: number;
}

interface NormalizedBounds {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

const boundsFor = (bounds: ChartPlotBounds): NormalizedBounds | null => {
  if (
    !Number.isFinite(bounds.x) ||
    !Number.isFinite(bounds.y) ||
    !Number.isFinite(bounds.width) ||
    !Number.isFinite(bounds.height) ||
    bounds.width < 0 ||
    bounds.height < 0
  ) {
    return null;
  }
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  if (!Number.isFinite(right) || !Number.isFinite(bottom)) return null;
  return { left: bounds.x, right, top: bounds.y, bottom };
};

const clampTo = (value: number, minimum: number, maximum: number): number =>
  Math.min(maximum, Math.max(minimum, value));

interface NormalizedAxis {
  readonly minimum: number;
  readonly maximum: number;
  readonly span: number;
  readonly scale: number;
  readonly sourceMinimum: number;
  readonly sourceMaximum: number;
}

/**
 * Build one common scale for every point on an axis. Scaling each endpoint by
 * its own magnitude changes the ratio of unequal extremes, which changes the
 * point where a connector crosses the plot. A shared per-axis scale preserves
 * that affine relationship while keeping arithmetic finite. The returned
 * bounds stay in normalized raw-data units; converting them to [0, 1] waits
 * until after clipping so a very narrow domain cannot amplify and cap points.
 */
const normalizedAxis = (
  samples: readonly ScatterLineSample[],
  domain: ScatterLineDomain,
  coordinate: 'x' | 'y',
): NormalizedAxis | null => {
  if (!Number.isFinite(domain.min) || !Number.isFinite(domain.max) || domain.min === domain.max) {
    return null;
  }
  let scale = Math.max(1, Math.abs(domain.min), Math.abs(domain.max));
  for (const sample of samples) {
    if (sample.kind !== 'point') continue;
    const value = sample[coordinate];
    if (Number.isFinite(value)) scale = Math.max(scale, Math.abs(value));
  }
  if (!Number.isFinite(scale) || scale <= 0) return null;
  let sourceMinimum = domain.min / scale;
  let sourceMaximum = domain.max / scale;
  if (sourceMinimum > sourceMaximum)
    [sourceMinimum, sourceMaximum] = [sourceMaximum, sourceMinimum];
  let span = sourceMaximum - sourceMinimum;
  if (!Number.isFinite(sourceMinimum) || !Number.isFinite(sourceMaximum)) return null;
  if (span === 0 || !Number.isFinite(span)) {
    // At sufficiently different magnitudes a narrow domain can collapse to a
    // single representable value after scaling. Keep clipping finite and treat
    // the collapsed boundary as a representable local interval; all source
    // points remain classified against the original domain below.
    const localStep = Number.EPSILON * Math.max(1, Math.abs(sourceMinimum));
    sourceMaximum = sourceMinimum + localStep;
    span = sourceMaximum - sourceMinimum;
    if (!(span > 0) || !Number.isFinite(span)) return null;
  }
  return {
    minimum: sourceMinimum,
    maximum: sourceMaximum,
    span,
    scale,
    sourceMinimum: Math.min(domain.min, domain.max),
    sourceMaximum: Math.max(domain.min, domain.max),
  };
};

/** Map one finite value into the shared normalized raw-data coordinate space. */
const normalizedCoordinate = (value: number, axis: NormalizedAxis): number => {
  if (!Number.isFinite(value)) return axis.minimum;
  const normalizedValue = value / axis.scale;
  if (Number.isFinite(normalizedValue)) return normalizedValue;
  return value < axis.sourceMinimum ? -Number.MAX_VALUE : Number.MAX_VALUE;
};

const projectNormalizedSegment = (
  segment: ScatterLineSegment,
  bounds: ChartPlotBounds,
  xAxis: NormalizedAxis,
  yAxis: NormalizedAxis,
  projection: ScatterLineProjection | undefined,
): ScatterLineSegment | null => {
  if (
    !Number.isFinite(bounds.x) ||
    !Number.isFinite(bounds.y) ||
    !Number.isFinite(bounds.width) ||
    !Number.isFinite(bounds.height) ||
    bounds.width < 0 ||
    bounds.height < 0
  ) {
    return null;
  }
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  if (!Number.isFinite(right) || !Number.isFinite(bottom)) return null;
  const x1Fraction = (segment.x1 - xAxis.minimum) / xAxis.span;
  const x2Fraction = (segment.x2 - xAxis.minimum) / xAxis.span;
  const y1Fraction = (segment.y1 - yAxis.minimum) / yAxis.span;
  const y2Fraction = (segment.y2 - yAxis.minimum) / yAxis.span;
  if (![x1Fraction, x2Fraction, y1Fraction, y2Fraction].every(Number.isFinite)) return null;
  const x1 =
    bounds.x + (projection?.reverseX === true ? 1 - x1Fraction : x1Fraction) * bounds.width;
  const x2 =
    bounds.x + (projection?.reverseX === true ? 1 - x2Fraction : x2Fraction) * bounds.width;
  const y1 =
    bounds.y + (projection?.reverseY === true ? y1Fraction : 1 - y1Fraction) * bounds.height;
  const y2 =
    bounds.y + (projection?.reverseY === true ? y2Fraction : 1 - y2Fraction) * bounds.height;
  if (![x1, x2, y1, y2].every(Number.isFinite)) return null;
  return {
    x1: clampTo(x1, bounds.x, right),
    y1: clampTo(y1, bounds.y, bottom),
    x2: clampTo(x2, bounds.x, right),
    y2: clampTo(y2, bounds.y, bottom),
    startIndex: segment.startIndex,
    endIndex: segment.endIndex,
  };
};

/**
 * Clip one finite segment to the plot rectangle.
 *
 * The endpoint coordinates are normalized before calculating the clipping
 * intersections. This keeps differences finite when a source value maps to a
 * very large screen coordinate (for example, opposing Number.MAX_VALUE data).
 */
export const clipScatterSegment = (
  first: ScatterLineSample & { readonly kind: 'point' },
  second: ScatterLineSample & { readonly kind: 'point' },
  bounds: ChartPlotBounds,
): ScatterLineSegment | null => {
  if (
    !Number.isFinite(first.x) ||
    !Number.isFinite(first.y) ||
    !Number.isFinite(second.x) ||
    !Number.isFinite(second.y)
  ) {
    return null;
  }
  const rectangle = boundsFor(bounds);
  if (rectangle === null) return null;

  const scale = Math.max(
    1,
    Math.abs(first.x),
    Math.abs(first.y),
    Math.abs(second.x),
    Math.abs(second.y),
    Math.abs(rectangle.left),
    Math.abs(rectangle.right),
    Math.abs(rectangle.top),
    Math.abs(rectangle.bottom),
  );
  let x1 = first.x / scale;
  let y1 = first.y / scale;
  let x2 = second.x / scale;
  let y2 = second.y / scale;
  const left = rectangle.left / scale;
  const right = rectangle.right / scale;
  const top = rectangle.top / scale;
  const bottom = rectangle.bottom / scale;

  const outcode = (x: number, y: number): number => {
    let code = 0;
    if (x < left) code |= 1;
    else if (x > right) code |= 2;
    if (y < top) code |= 4;
    else if (y > bottom) code |= 8;
    return code;
  };

  // Cohen–Sutherland keeps the clipped coordinate on the active boundary
  // explicitly. That matters when a Number.MAX_VALUE endpoint makes a plot
  // bound vanish under normalization; multiplying the normalized result back
  // would otherwise turn the boundary into zero and create a false point.
  let code1 = outcode(x1, y1);
  let code2 = outcode(x2, y2);
  let x1Screen = first.x;
  let y1Screen = first.y;
  let x2Screen = second.x;
  let y2Screen = second.y;
  let iterations = 0;
  while (iterations < 16) {
    iterations += 1;
    if ((code1 | code2) === 0) break;
    if ((code1 & code2) !== 0) return null;
    const outside = code1 !== 0 ? code1 : code2;
    let intersectionX = 0;
    let intersectionY = 0;
    let boundaryX: number | undefined;
    let boundaryY: number | undefined;
    const dx = x2 - x1;
    const dy = y2 - y1;
    if ((outside & 8) !== 0) {
      if (dy === 0) return null;
      intersectionY = bottom;
      intersectionX = x1 + (dx * (bottom - y1)) / dy;
      boundaryY = rectangle.bottom;
    } else if ((outside & 4) !== 0) {
      if (dy === 0) return null;
      intersectionY = top;
      intersectionX = x1 + (dx * (top - y1)) / dy;
      boundaryY = rectangle.top;
    } else if ((outside & 2) !== 0) {
      if (dx === 0) return null;
      intersectionX = right;
      intersectionY = y1 + (dy * (right - x1)) / dx;
      boundaryX = rectangle.right;
    } else {
      if (dx === 0) return null;
      intersectionX = left;
      intersectionY = y1 + (dy * (left - x1)) / dx;
      boundaryX = rectangle.left;
    }
    if (!Number.isFinite(intersectionX) || !Number.isFinite(intersectionY)) return null;
    const screenX = boundaryX ?? intersectionX * scale;
    const screenY = boundaryY ?? intersectionY * scale;
    if (!Number.isFinite(screenX) || !Number.isFinite(screenY)) return null;
    if (outside === code1) {
      // Snap only the active boundary. The other coordinate may still be
      // outside the rectangle when the endpoint is at a corner; clamping it
      // here invents an inside point and can make a same-side segment appear
      // to cross the plot.
      x1 = boundaryX === undefined ? intersectionX : boundaryX / scale;
      y1 = boundaryY === undefined ? intersectionY : boundaryY / scale;
      x1Screen = screenX;
      y1Screen = screenY;
      code1 = outcode(x1, y1);
    } else {
      x2Screen = screenX;
      y2Screen = screenY;
      // Keep the normalized endpoint in lockstep with the active boundary so
      // a second clipping pass (a corner crossing) remains stable, while the
      // inactive coordinate retains its true intersection value.
      x2 = boundaryX === undefined ? intersectionX : boundaryX / scale;
      y2 = boundaryY === undefined ? intersectionY : boundaryY / scale;
      code2 = outcode(x2, y2);
    }
  }
  if ((code1 | code2) !== 0) return null;

  // The result is clamped as a final defense against one-ulp interpolation
  // drift.
  const finalX1 = clampTo(x1Screen, rectangle.left, rectangle.right);
  const finalY1 = clampTo(y1Screen, rectangle.top, rectangle.bottom);
  const finalX2 = clampTo(x2Screen, rectangle.left, rectangle.right);
  const finalY2 = clampTo(y2Screen, rectangle.top, rectangle.bottom);
  if (
    !Number.isFinite(finalX1) ||
    !Number.isFinite(finalY1) ||
    !Number.isFinite(finalX2) ||
    !Number.isFinite(finalY2)
  ) {
    return null;
  }
  return {
    x1: finalX1,
    y1: finalY1,
    x2: finalX2,
    y2: finalY2,
    startIndex: first.sourceIndex,
    endIndex: second.sourceIndex,
  };
};

/**
 * Connect adjacent finite samples in source order and clip each segment to the
 * plot rectangle. Blanks break a run unless `spanBlanks` is enabled; explicit
 * breaks always end a run, including in span mode.
 */
export const scatterLineSegments = (
  samples: readonly ScatterLineSample[],
  bounds: ChartPlotBounds,
  spanBlanks = false,
  domains?: ScatterLineDomains,
  projection?: ScatterLineProjection,
): readonly ScatterLineSegment[] => {
  const segments: ScatterLineSegment[] = [];
  const xAxis = domains === undefined ? null : normalizedAxis(samples, domains.x, 'x');
  const yAxis = domains === undefined ? null : normalizedAxis(samples, domains.y, 'y');
  if (domains !== undefined && (xAxis === null || yAxis === null)) return segments;
  let previous: (ScatterLineSample & { readonly kind: 'point' }) | null = null;
  for (const sample of samples) {
    if (sample.kind === 'break') {
      previous = null;
      continue;
    }
    if (sample.kind === 'blank') {
      if (!spanBlanks) previous = null;
      continue;
    }
    if (previous !== null) {
      const first =
        domains === undefined
          ? previous
          : {
              kind: 'point' as const,
              x: normalizedCoordinate(previous.x, xAxis as NormalizedAxis),
              y: normalizedCoordinate(previous.y, yAxis as NormalizedAxis),
              sourceIndex: previous.sourceIndex,
            };
      const second =
        domains === undefined
          ? sample
          : {
              kind: 'point' as const,
              x: normalizedCoordinate(sample.x, xAxis as NormalizedAxis),
              y: normalizedCoordinate(sample.y, yAxis as NormalizedAxis),
              sourceIndex: sample.sourceIndex,
            };
      const segment = clipScatterSegment(
        first,
        second,
        domains === undefined
          ? bounds
          : {
              x: xAxis?.minimum ?? 0,
              y: yAxis?.minimum ?? 0,
              width: xAxis?.span ?? 0,
              height: yAxis?.span ?? 0,
            },
      );
      if (segment !== null) {
        const projected =
          domains === undefined
            ? segment
            : projectNormalizedSegment(
                segment,
                bounds,
                xAxis as NormalizedAxis,
                yAxis as NormalizedAxis,
                projection,
              );
        if (projected !== null) segments.push(projected);
      }
    }
    previous = sample;
  }
  return segments;
};
