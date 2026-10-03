// Formula-text model for the Mac formula palette: reads and writes the
// single outer call `=NAME(arg, ...)` the palette's argument fields project.

import { splitFormulaArgsAllowEmpty } from '../render/conditional-formula/parser.js';

export interface ParsedFormulaCall {
  name: string;
  args: string[];
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
