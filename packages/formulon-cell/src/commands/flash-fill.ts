/**
 * Flash Fill (Ctrl+E) — pattern inference from a small set of
 * (input, output) examples in adjacent columns. The user supplies one or
 * more examples in the target column; Flash Fill extends the pattern to
 * the rest of the column.
 *
 * the canonical implementation uses a learned program-synthesis model
 * (FlashFill / Prose). The algorithm here is a deliberately small
 * heuristic that covers the common cases users actually run Ctrl+E for:
 *
 *   1. Identity         — output equals input.
 *   2. Constant suffix  — output = input + literal.
 *   3. Constant prefix  — output = literal + input.
 *   4. Affix wrap       — output = prefix + input + suffix.
 *   5. Substring slice  — output is a contiguous slice of the input.
 *   6. Delimiter token  — output is the n-th split-on-delimiter token.
 *   7. Casing transform — output is upper / lower / title case of the input.
 *
 * The module exports `inferFlashFillPattern` (pure) and `applyFlashFill`
 * (apply an inferred pattern to a row of pending inputs). Both are unit-
 * tested in `tests/unit/commands/flash-fill.test.ts`.
 */

export type FlashFillPattern =
  | { kind: 'identity' }
  | { kind: 'constant-suffix'; suffix: string }
  | { kind: 'constant-prefix'; prefix: string }
  | { kind: 'affix'; prefix: string; suffix: string }
  | { kind: 'substring'; start: number; length: number }
  | { kind: 'token'; delimiter: string; index: number }
  | { kind: 'case'; mode: 'upper' | 'lower' | 'title' };

export interface FlashFillExample {
  input: string;
  output: string;
}

const TITLE_CASE = (s: string): string =>
  s.replace(/\w\S*/g, (w) =>
    w.length > 0 ? (w[0] ?? '').toUpperCase() + w.slice(1).toLowerCase() : w,
  );

/** Apply a single pattern to one input. Returns null when the pattern
 *  cannot be applied (for example, slice bounds beyond the input). */
export function applyFlashFillPattern(pattern: FlashFillPattern, input: string): string | null {
  switch (pattern.kind) {
    case 'identity':
      return input;
    case 'constant-suffix':
      return input + pattern.suffix;
    case 'constant-prefix':
      return pattern.prefix + input;
    case 'affix':
      return pattern.prefix + input + pattern.suffix;
    case 'substring': {
      if (pattern.start < 0) return null;
      if (pattern.start > input.length) return null;
      return input.substring(pattern.start, pattern.start + pattern.length);
    }
    case 'token': {
      const tokens = input.split(pattern.delimiter);
      return tokens[pattern.index] ?? null;
    }
    case 'case':
      if (pattern.mode === 'upper') return input.toUpperCase();
      if (pattern.mode === 'lower') return input.toLowerCase();
      return TITLE_CASE(input);
  }
}

/** True when `pattern` reproduces every example. Used during inference to
 *  validate a candidate against the rest of the example set. */
function patternMatchesAll(
  pattern: FlashFillPattern,
  examples: readonly FlashFillExample[],
): boolean {
  for (const ex of examples) {
    if (applyFlashFillPattern(pattern, ex.input) !== ex.output) return false;
  }
  return true;
}

/** Try to infer a single pattern that explains every example. Returns the
 *  first matching candidate in the order listed at the top of this file —
 *  the simpler patterns are preferred. Returns null when nothing fits. */
export function inferFlashFillPattern(
  examples: readonly FlashFillExample[],
): FlashFillPattern | null {
  if (examples.length === 0) return null;
  const first = examples[0];
  if (!first) return null;

  // 1. Identity.
  if (patternMatchesAll({ kind: 'identity' }, examples)) {
    return { kind: 'identity' };
  }

  // 2. Constant suffix — output starts with input.
  if (first.output.startsWith(first.input) && first.input.length > 0) {
    const suffix = first.output.slice(first.input.length);
    const cand: FlashFillPattern = { kind: 'constant-suffix', suffix };
    if (patternMatchesAll(cand, examples)) return cand;
  }

  // 3. Constant prefix — output ends with input.
  if (first.output.endsWith(first.input) && first.input.length > 0) {
    const prefix = first.output.slice(0, first.output.length - first.input.length);
    const cand: FlashFillPattern = { kind: 'constant-prefix', prefix };
    if (patternMatchesAll(cand, examples)) return cand;
  }

  // 4. Affix — output contains input bracketed by literal text.
  if (first.input.length > 0) {
    const idx = first.output.indexOf(first.input);
    if (idx >= 0) {
      const prefix = first.output.slice(0, idx);
      const suffix = first.output.slice(idx + first.input.length);
      const cand: FlashFillPattern = { kind: 'affix', prefix, suffix };
      if (patternMatchesAll(cand, examples)) return cand;
    }
  }

  // 5. Delimiter token — split on common separators and pick the slot
  //    whose value matches the output for every example. Tried before raw
  //    substring slicing because "first word" is more semantically
  //    portable than "first N characters" — `John Smith → John` and
  //    `Jane Doe → Jane` happen to match both, but `Jane Williams → Jane`
  //    only matches the token rule.
  for (const delim of [' ', ',', '@', '/', '-', '\t', '_']) {
    const tokens0 = first.input.split(delim);
    const idx = tokens0.indexOf(first.output);
    if (idx < 0) continue;
    const cand: FlashFillPattern = { kind: 'token', delimiter: delim, index: idx };
    if (patternMatchesAll(cand, examples)) return cand;
  }

  // 6. Substring slice — output is a contiguous chunk of input.
  if (first.input.includes(first.output) && first.output.length > 0) {
    const start = first.input.indexOf(first.output);
    const cand: FlashFillPattern = {
      kind: 'substring',
      start,
      length: first.output.length,
    };
    if (patternMatchesAll(cand, examples)) return cand;
  }

  // 7. Casing transforms.
  for (const mode of ['upper', 'lower', 'title'] as const) {
    const cand: FlashFillPattern = { kind: 'case', mode };
    if (patternMatchesAll(cand, examples)) return cand;
  }

  return null;
}

/** Apply an inferred pattern across a sequence of pending inputs. Inputs
 *  that fail the pattern produce `null`; callers usually fall back to
 *  leaving those cells untouched. */
export function applyFlashFill(
  pattern: FlashFillPattern,
  inputs: readonly string[],
): (string | null)[] {
  return inputs.map((s) => applyFlashFillPattern(pattern, s));
}
