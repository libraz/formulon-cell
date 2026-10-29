/** Lexical context at a formula caret, shared by reference and autocomplete
 * insertion. Quotes and structured-reference headers are not formula tokens. */
export type FormulaCaretContext = 'expression' | 'string' | 'sheet' | 'bracket' | 'function';

export function formulaCaretContext(text: string, caret: number): FormulaCaretContext {
  let i = text.startsWith('=') ? 1 : 0;
  const stop = Math.max(i, Math.min(caret, text.length));
  let inString = false;
  let inSheet = false;
  let bracketDepth = 0;
  while (i < stop) {
    const ch = text[i] ?? '';
    const next = text[i + 1] ?? '';
    if (inString) {
      if (ch === '"') {
        if (next === '"') {
          i += 2;
          continue;
        }
        inString = false;
      }
      i += 1;
      continue;
    }
    if (inSheet) {
      if (ch === "'") {
        if (next === "'") {
          i += 2;
          continue;
        }
        inSheet = false;
      }
      i += 1;
      continue;
    }
    if (bracketDepth > 0) {
      if (ch === "'" && ['[', ']', '@', '#', "'"].includes(next)) {
        i += 2;
        continue;
      }
      if (ch === '[') bracketDepth += 1;
      else if (ch === ']') bracketDepth -= 1;
      i += 1;
      continue;
    }
    if (ch === '"') {
      inString = true;
      i += 1;
      continue;
    }
    if (ch === "'") {
      // A quoted sheet qualifier starts at a token boundary and closes before
      // the qualifier's `!`; apostrophes inside table headers are handled above.
      inSheet = true;
      i += 1;
      continue;
    }
    if (ch === '[') {
      // R1C1 relative coordinates use brackets as part of a reference atom.
      const before = text.slice(0, i);
      if (/(?:^|[^\p{L}\p{N}_.])R(?:\[-?\d+\])?C?(?:\[-?\d*)?$/iu.test(before)) {
        i += 1;
        continue;
      }
      bracketDepth = 1;
      i += 1;
      continue;
    }
    i += 1;
  }
  if (inString) return 'string';
  if (inSheet) return 'sheet';
  if (bracketDepth > 0) return 'bracket';
  // A caret inside an unquoted sheet name (including before its `!`) should
  // not offer function/name completions.
  const after = text.slice(stop);
  const before = text.slice(0, stop);
  const wordBefore = /[\p{L}\p{N}_.]+$/u.exec(before)?.[0] ?? '';
  const wordAfter = /^[\p{L}\p{N}_.]+/u.exec(after)?.[0] ?? '';
  const wordStart = before.length - wordBefore.length;
  const preceding = before[wordStart - 1] ?? '';
  if (
    (wordBefore || wordAfter) &&
    !IDENT_CHAR.test(preceding) &&
    after.slice(wordAfter.length).startsWith('!')
  )
    return 'sheet';
  if (
    (wordBefore || wordAfter) &&
    !IDENT_CHAR.test(preceding) &&
    after.slice(wordAfter.length).startsWith('(')
  )
    return 'function';
  return 'expression';
}

const IDENT_CHAR = /[\p{L}\p{N}_.\\]/u;
