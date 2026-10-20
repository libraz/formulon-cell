import { describe, expect, it } from 'vitest';
import {
  automaticAxisDomain,
  automaticMajorUnit,
  linearScale,
  majorTicks,
  niceDomain,
  niceTicks,
  numericDomain,
  powerTicks,
  resolvePositivePowerDomain,
  unclampedLogProjection,
  validateLogBase,
} from '../../../src/charts/scales.js';

describe('chart scales', () => {
  it('picks the automatic major unit Excel picks', () => {
    expect(automaticMajorUnit(15.75)).toBe(2);
    expect(automaticMajorUnit(105)).toBe(20);
    expect(automaticMajorUnit(7.35)).toBe(1);
    expect(automaticMajorUnit(1.05)).toBeCloseTo(0.2);
    expect(automaticMajorUnit(5)).toBe(1);
  });

  it('bounds an automatic axis the way Excel does', () => {
    const bounds = (min: number, max: number, fixed?: { min?: number; max?: number }) => {
      const domain = automaticAxisDomain(min, max, fixed);
      return [domain.min, domain.max, domain.majorUnit];
    };
    // A twentieth of headroom, rounded out to the unit.
    expect(bounds(3, 15)).toEqual([0, 16, 2]);
    expect(bounds(0, 100)).toEqual([0, 120, 20]);
    expect(bounds(0, 10)).toEqual([0, 12, 2]);
    // A single positive value starts at zero.
    expect(bounds(5, 5)).toEqual([0, 6, 1]);
    // Data spanning less than a sixth of its maximum starts below the minimum.
    expect(bounds(110, 120)).toEqual([104, 122, 2]);
    // Negative data mirrors positive data.
    expect(bounds(-15, -3)).toEqual([-16, 0, 2]);
    // Data either side of zero gets headroom on both sides.
    expect(bounds(-10, 10)).toEqual([-12, 12, 2]);
    // A fixed side keeps its value.
    expect(bounds(3, 15, { max: 20 })).toEqual([0, 20, 2]);
    expect(bounds(3, 15, { min: 2 })).toEqual([2, 16, 2]);
  });

  it('keeps extreme automatic axes finite', () => {
    for (const [min, max] of [
      [0, Number.MAX_VALUE],
      [-Number.MAX_VALUE, Number.MAX_VALUE],
      [Number.MIN_VALUE, Number.MIN_VALUE],
    ] as const) {
      const domain = automaticAxisDomain(min, max);
      expect(Number.isFinite(domain.min)).toBe(true);
      expect(Number.isFinite(domain.max)).toBe(true);
      expect(domain.max).toBeGreaterThan(domain.min);
      expect(domain.majorUnit).toBeGreaterThan(0);
    }
  });

  it('keeps explicit tick bounds instead of expanding them', () => {
    const ticks = niceTicks(1, 9, 5);
    expect(ticks.length).toBeGreaterThan(0);
    expect(ticks.every((tick) => tick >= 1 && tick <= 9)).toBe(true);
    expect(ticks).not.toContain(0);
    expect(ticks).not.toContain(10);
  });

  it('preserves one-sided bounds while synthesizing the missing side', () => {
    const lower = niceDomain(10, 5, 5);
    expect(lower.min).toBe(5);
    const upper = niceDomain(5, 10, 5);
    expect(upper.max).toBe(10);
    expect(lower.max).toBeGreaterThan(lower.min);
    expect(upper.min).toBeLessThan(upper.max);
  });

  it('maps extreme and tiny finite values without NaN or Infinity', () => {
    const values = [
      [Number.MIN_VALUE, Number.MIN_VALUE * 2, Number.MIN_VALUE * 3],
      [1e-13, 2e-13, 3e-13],
      [-1e308, 0, 1e308],
      [Number.MAX_VALUE * 0.5, Number.MAX_VALUE * 0.75, Number.MAX_VALUE],
    ] as const;
    for (const domain of values) {
      const mapped = domain.map((value) => linearScale(value, domain[0], domain[2], 20, 112));
      expect(mapped.every((value) => value !== null && Number.isFinite(value))).toBe(true);
    }
    expect(numericDomain([Number.MAX_VALUE])).toEqual({
      min: expect.any(Number),
      max: expect.any(Number),
    });
    expect(niceTicks(Number.MIN_VALUE, Number.MIN_VALUE * 3).every(Number.isFinite)).toBe(true);
    expect(niceTicks(-1e308, 1e308).every(Number.isFinite)).toBe(true);
  });

  it('starts explicit major units at the domain minimum and thins by integer multiples', () => {
    expect(majorTicks(3, 13, 5)).toEqual([3, 8, 13]);
    const dense = majorTicks(0, 1000, 1);
    expect(dense).toHaveLength(101);
    expect(dense[0]).toBe(0);
    expect(dense.at(-1)).toBe(1000);
    expect(
      dense.every((tick, index) => index === 0 || (tick - (dense[index - 1] ?? 0)) % 10 === 0),
    ).toBe(true);
    expect(majorTicks(0, Number.MIN_VALUE * 3, Number.MIN_VALUE)).toEqual([
      0,
      Number.MIN_VALUE,
      Number.MIN_VALUE * 2,
      Number.MIN_VALUE * 3,
    ]);
    const fractional = majorTicks(0, 0.3, 0.1);
    expect(fractional).toHaveLength(4);
    expect(fractional[0]).toBe(0);
    expect(fractional[1]).toBeCloseTo(0.1);
    expect(fractional[2]).toBeCloseTo(0.2);
    expect(fractional[3]).toBe(0.3);
    expect(fractional.every((tick) => tick >= 0 && tick <= 0.3)).toBe(true);
  });

  it('rejects invalid major units and remains finite for extreme ranges', () => {
    for (const unit of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
      const ticks = majorTicks(-1, 1, unit);
      expect(ticks.length).toBeGreaterThan(0);
      expect(ticks.length).toBeLessThanOrEqual(101);
      expect(ticks.every(Number.isFinite)).toBe(true);
    }
    const extreme = majorTicks(-Number.MAX_VALUE, Number.MAX_VALUE, 1);
    expect(extreme.length).toBeGreaterThan(0);
    expect(extreme.length).toBeLessThanOrEqual(101);
    expect(extreme.every(Number.isFinite)).toBe(true);
  });

  it('resolves finite positive log domains without admitting zero or invalid bounds', () => {
    expect(validateLogBase(2)).toBe(2);
    expect(validateLogBase(1000)).toBe(1000);
    expect(validateLogBase(1.99)).toBeUndefined();
    expect(validateLogBase(Number.NaN)).toBeUndefined();
    expect(
      resolvePositivePowerDomain([0, -2, Number.NaN, 2, 80], undefined, undefined, 10),
    ).toEqual({ min: 1, max: 100 });
    expect(resolvePositivePowerDomain([], undefined, undefined, 10)).toEqual({ min: 1, max: 10 });
    expect(resolvePositivePowerDomain([2, 80], 100, 1, 10)).toEqual({ min: 1, max: 100 });
    expect(resolvePositivePowerDomain([10], undefined, undefined, 10)).toEqual({
      min: 1,
      max: 100,
    });
    expect(resolvePositivePowerDomain([2], undefined, undefined, 10)).toEqual({
      min: 1,
      max: 10,
    });
    expect(resolvePositivePowerDomain([10 - Number.EPSILON * 8], undefined, undefined, 10)).toEqual(
      {
        min: 1,
        max: 10,
      },
    );
    expect(resolvePositivePowerDomain([10 + Number.EPSILON * 8], undefined, undefined, 10)).toEqual(
      {
        min: 10,
        max: 100,
      },
    );
    expect(resolvePositivePowerDomain([1000, 10000], undefined, undefined, 10)).toEqual({
      min: 1000,
      max: 10000,
    });
    expect(resolvePositivePowerDomain([2], -1, Number.POSITIVE_INFINITY, 10)).toEqual({
      min: 1,
      max: 10,
    });
    expect(resolvePositivePowerDomain([1], 100, undefined, 10)).toEqual({ min: 100, max: 1000 });
    expect(resolvePositivePowerDomain([], 1000, undefined, 10)).toEqual({ min: 1000, max: 10000 });
    expect(resolvePositivePowerDomain([1000], undefined, 1, 10)).toEqual({ min: 0.1, max: 1 });
    expect(resolvePositivePowerDomain([1], 1000, undefined, 10)).toEqual({
      min: 1000,
      max: 10000,
    });
    expect(resolvePositivePowerDomain([], 100, undefined, 10)).toEqual({ min: 100, max: 1000 });
    expect(resolvePositivePowerDomain([], undefined, 1, 10)).toEqual({ min: 0.1, max: 1 });
    const maximumBound = resolvePositivePowerDomain([], Number.MAX_VALUE, undefined, 10);
    expect(maximumBound.min).toBeGreaterThan(0);
    expect(maximumBound.min).toBeLessThan(maximumBound.max);
    expect(maximumBound.max).toBe(Number.MAX_VALUE);
    const minimumBound = resolvePositivePowerDomain([], undefined, Number.MIN_VALUE, 10);
    expect(minimumBound.min).toBe(Number.MIN_VALUE);
    expect(minimumBound.max).toBeGreaterThan(minimumBound.min);
  });

  it('generates bounded ascending power ticks and keeps log projection unclamped', () => {
    const ticks = powerTicks(1, Number.MAX_VALUE, 2);
    expect(ticks.length).toBeLessThanOrEqual(101);
    expect(ticks.every(Number.isFinite)).toBe(true);
    expect(ticks.every((tick, index) => index === 0 || tick > (ticks[index - 1] as number))).toBe(
      true,
    );
    expect(powerTicks(1, 100, 10)).toEqual([1, 10, 100]);
    expect(powerTicks(2, 3, 10)).toEqual([]);
    expect(unclampedLogProjection(1, 1, 100, 20, 220)).toBe(20);
    expect(unclampedLogProjection(100, 1, 100, 20, 220)).toBe(220);
    expect(unclampedLogProjection(10, 1, 100, 20, 220)).toBeCloseTo(120);
    expect(
      Number.isFinite(
        unclampedLogProjection(Number.MAX_VALUE, Number.MAX_VALUE / 2, Number.MAX_VALUE, 20, 220),
      ),
    ).toBe(true);
  });
});
