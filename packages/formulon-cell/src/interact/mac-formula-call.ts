// Formula-text model for the Mac formula palette: reads and writes the
// single outer call `=NAME(arg, ...)` the palette's argument fields project.

import { type FunctionCallSpan, findFunctionCallAtCaret } from '../commands/refs.js';
import { splitFormulaArgsAllowEmpty } from '../render/conditional-formula/parser.js';

export interface ParsedFormulaCall {
  name: string;
  args: string[];
}

/** A live scanner projection of the function call under an editor caret. */
export interface ProjectedFormulaCall {
  readonly source: string;
  readonly span: FunctionCallSpan;
  readonly args: readonly string[];
}

export const canonicalName = (name: string): string => name.trim().toLocaleUpperCase();

const listSeparator = (value: string | undefined): ',' | ';' => (value === ';' ? ';' : ',');

/** The shared conditional-formula splitter is comma based. Palette drafts
 * need the workbook's active separator while retaining commas in quoted,
 * array, and structured-reference arguments, so keep this small display-side
 * splitter local to the palette call model. */
const splitPaletteArgs = (body: string, separator: ',' | ';'): string[] | null => {
  if (separator === ',') return splitFormulaArgsAllowEmpty(body);
  const args: string[] = [];
  let start = 0;
  let quote: '"' | "'" | null = null;
  const stack: string[] = [];
  for (let i = 0; i < body.length; i += 1) {
    const ch = body[i] ?? '';
    const next = body[i + 1] ?? '';
    if (quote) {
      if (ch === quote) {
        if (next === quote) {
          i += 1;
        } else {
          quote = null;
        }
      }
      continue;
    }
    const top = stack[stack.length - 1];
    if (top === '[') {
      if (ch === '[') {
        stack.push(ch);
        continue;
      }
      if (ch === ']') {
        stack.pop();
        continue;
      }
      // Structured references use apostrophes to escape [, ], #, @, and
      // apostrophe itself inside a header token rather than opening a sheet
      // name quote. Match the shared formula splitter's lexical mode.
      if (ch === "'" && "[]#'@".includes(next)) {
        i += 1;
      }
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(' || ch === '[' || ch === '{') {
      stack.push(ch);
      continue;
    }
    if (ch === ')' || ch === ']' || ch === '}') {
      const expected = ch === ')' ? '(' : ch === ']' ? '[' : '{';
      if (stack.pop() !== expected) return null;
      continue;
    }
    if (ch === separator && stack.length === 0) {
      args.push(body.slice(start, i));
      start = i + 1;
    }
  }
  if (quote || stack.length > 0) return null;
  args.push(body.slice(start));
  return args;
};

export const parseOuterCall = (
  raw: string,
  expectedName: string,
  activeListSeparator?: string,
  resolveKnownName?: (rawName: string) => string | null,
): ParsedFormulaCall | null => {
  const match = raw.trim().match(/^=\s*([\p{L}_][\p{L}\p{N}_.]*)\s*\(([\s\S]*)\)\s*$/u);
  if (!match) return null;
  const rawName = match[1] ?? '';
  const resolvedName = resolveKnownName?.(rawName) ?? canonicalName(rawName);
  if (!resolvedName || canonicalName(resolvedName) !== canonicalName(expectedName)) return null;
  const body = match[2] ?? '';
  const args = splitPaletteArgs(body, listSeparator(activeListSeparator));
  return args === null ? null : { name: canonicalName(rawName), args };
};

const trailingBlankIndex = (args: readonly string[]): number => {
  for (let i = args.length - 1; i >= 0; i -= 1) {
    if ((args[i] ?? '').trim() !== '') return i;
  }
  return -1;
};

export const assembledFormula = (
  name: string,
  args: readonly string[],
  activeListSeparator?: string,
): string => {
  const end = trailingBlankIndex(args);
  const separator = listSeparator(activeListSeparator);
  return `=${name}(${end < 0 ? '' : args.slice(0, end + 1).join(separator)})`;
};

/** `=NAME(...)` keeping exactly `count` arguments, or trimming trailing blanks when `count` is null. */
export const formulaWithArgumentCount = (
  name: string,
  args: readonly string[],
  count: number | null,
  activeListSeparator?: string,
): string => {
  const separator = listSeparator(activeListSeparator);
  return count === null
    ? assembledFormula(name, args, separator)
    : `=${name}(${args.slice(0, count).join(separator)})`;
};

/**
 * Project the live, known function call containing `caret` without parsing or
 * rewriting its surrounding formula. The scanner owns all delimiter and quote
 * handling; argument slices stay verbatim so editing one field cannot rewrite
 * whitespace or other text in its untouched siblings.
 */
export const projectFormulaCallAtCaret = (
  raw: string,
  caret: number,
  resolveKnownName: (rawName: string) => string | null,
  activeListSeparator?: string,
): ProjectedFormulaCall | null => {
  const span = findFunctionCallAtCaret(raw, caret, resolveKnownName, activeListSeparator);
  if (!span || (!span.complete && !span.safeToComplete)) return null;

  return {
    source: raw,
    span,
    args: span.argumentSpans.map(({ start, end }) => raw.slice(start, end)),
  };
};

/**
 * Replace only the projected call after confirming that its source is still
 * current. The caller supplies the desired function arguments; the returned
 * caret sits at the one-past-call boundary recognised by the scanner.
 */
export const replaceProjectedFormulaCall = (
  currentRaw: string,
  projection: ProjectedFormulaCall,
  name: string,
  args: readonly string[],
  explicitArgumentCount: number | null,
  activeListSeparator?: string,
): { raw: string; caret: number } | null => {
  if (currentRaw !== projection.source) return null;

  const replacement = formulaWithArgumentCount(
    name,
    args,
    explicitArgumentCount,
    activeListSeparator,
  ).slice(1);
  const { start, end } = projection.span.call;
  return {
    raw: `${currentRaw.slice(0, start)}${replacement}${currentRaw.slice(end)}`,
    caret: start + replacement.length,
  };
};
