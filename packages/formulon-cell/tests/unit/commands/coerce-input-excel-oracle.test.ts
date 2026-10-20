import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { coerceInput } from '../../../src/commands/coerce-input.js';
import { formatCodeToNumFmt } from '../../../src/engine/format-writeback.js';
import { excelSerialFromParts, excelSerialToDateParts } from '../../../src/format/date-serial.js';

interface OracleCase {
  input: string;
  kind: 'number' | 'text' | 'bool' | 'error';
  value: number | string | boolean;
  format?: string;
  /** The input omits the year, so Excel filled in the year it was captured in. */
  currentYear?: true;
}

interface Oracle {
  captureYear: number;
  locales: Record<string, OracleCase[]>;
}

// Captured from desktop Excel by scripts/excel-input-probe.py.
const oracle = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-input-coercion.json'), 'utf8'),
) as Oracle;

// A typed leading `+`, `-` or `@` starts a formula in the desktop UI, but the
// object model the probe drives stores the same text verbatim, so the capture
// does not show typed behaviour for these.
const NOT_OBSERVABLE = new Set(['+1+1', '-1+1', '@SUM(1)']);

const expectedNumber = (c: OracleCase): number => {
  const value = c.value as number;
  if (!c.currentYear) return value;
  const parts = excelSerialToDateParts(Math.floor(value));
  if (!parts) throw new Error(`bad serial ${value}`);
  const year = new Date().getUTCFullYear();
  const day = excelSerialFromParts(year, parts.month, parts.day);
  if (day === null) throw new Error(`bad date ${year}-${parts.month}-${parts.day}`);
  return day + (value - Math.floor(value));
};

describe('coerceInput matches desktop Excel typed-input parsing', () => {
  for (const [locale, cases] of Object.entries(oracle.locales)) {
    describe(locale, () => {
      for (const c of cases) {
        if (NOT_OBSERVABLE.has(c.input)) continue;
        it(JSON.stringify(c.input), () => {
          const got = coerceInput(c.input, { locale });
          if (c.input.startsWith('=')) {
            expect(got.kind).toBe('formula');
            return;
          }
          switch (c.kind) {
            case 'number': {
              expect(got.kind).toBe('number');
              if (got.kind !== 'number') return;
              const want = expectedNumber(c);
              expect(Math.abs(got.value - want)).toBeLessThanOrEqual(
                1e-9 * Math.max(1, Math.abs(want)),
              );
              expect(got.implicitFormat ?? null).toEqual(
                c.format ? formatCodeToNumFmt(c.format) : null,
              );
              return;
            }
            case 'bool':
              expect(got).toEqual({ kind: 'bool', value: c.value });
              return;
            case 'error':
              expect(got).toMatchObject({ kind: 'error', text: c.value });
              return;
            default:
              expect(got).toEqual({
                kind: 'text',
                value: c.input.startsWith("'") ? c.value : c.input,
              });
          }
        });
      }
    });
  }
});
