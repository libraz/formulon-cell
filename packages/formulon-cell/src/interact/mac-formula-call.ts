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

export const canonicalName = (name: string): string => name.trim().toUpperCase();

export const parseOuterCall = (raw: string, expectedName: string): ParsedFormulaCall | null => {
  const match = raw.trim().match(/^=\s*([A-Za-z][A-Za-z0-9.]*)\s*\(([\s\S]*)\)\s*$/);
  if (!match || canonicalName(match[1] ?? '') !== canonicalName(expectedName)) return null;
  const body = match[2] ?? '';
  const args = splitFormulaArgsAllowEmpty(body);
  return args === null ? null : { name: canonicalName(match[1] ?? ''), args };
};

const trailingBlankIndex = (args: readonly string[]): number => {
  for (let i = args.length - 1; i >= 0; i -= 1) {
    if ((args[i] ?? '').trim() !== '') return i;
  }
  return -1;
};

export const assembledFormula = (name: string, args: readonly string[]): string => {
  const end = trailingBlankIndex(args);
  return `=${name}(${end < 0 ? '' : args.slice(0, end + 1).join(',')})`;
};

/** `=NAME(...)` keeping exactly `count` arguments, or trimming trailing blanks when `count` is null. */
export const formulaWithArgumentCount = (
  name: string,
  args: readonly string[],
  count: number | null,
): string =>
  count === null ? assembledFormula(name, args) : `=${name}(${args.slice(0, count).join(',')})`;

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
): ProjectedFormulaCall | null => {
  const span = findFunctionCallAtCaret(raw, caret, resolveKnownName);
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
): { raw: string; caret: number } | null => {
  if (currentRaw !== projection.source) return null;

  const replacement = formulaWithArgumentCount(name, args, explicitArgumentCount).slice(1);
  const { start, end } = projection.span.call;
  return {
    raw: `${currentRaw.slice(0, start)}${replacement}${currentRaw.slice(end)}`,
    caret: start + replacement.length,
  };
};
