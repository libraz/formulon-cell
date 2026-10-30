import { colFromLetters, MAX_COL, MAX_ROW } from '../engine/address.js';
import { formulaCaretContext } from './formula-edit-context.js';
import { rewriteRefs } from './formula-ref-lexer.js';

export interface FormulaReferenceInputOptions {
  /** The current editor displays native relative / absolute R1C1 references. */
  r1c1?: boolean;
  /** Reject a non-collapsed selection that touches string text or starts or
   *  ends inside a structured-reference header (keyboard point navigation). */
  requireExpressionSelection?: boolean;
}

export interface FormulaReferenceInsertionSpan {
  start: number;
  end: number;
}

const IDENT_CHAR = /[\p{L}\p{N}_.\\]/u;
const R1C1_ATOM = /R(?:\[-?\d+\]|[1-9]\d*)?C(?:\[-?\d+\]|[1-9]\d*)?/iy;

function mergeColonConnectedSpans(
  text: string,
  spans: FormulaReferenceInsertionSpan[],
): FormulaReferenceInsertionSpan[] {
  const merged: FormulaReferenceInsertionSpan[] = [];
  for (const span of spans) {
    const previous = merged[merged.length - 1];
    if (previous && /^\s*:\s*$/u.test(text.slice(previous.end, span.start))) {
      previous.end = span.end;
    } else {
      merged.push({ ...span });
    }
  }
  return merged;
}

function a1ReferenceSpans(text: string): FormulaReferenceInsertionSpan[] {
  const spans: FormulaReferenceInsertionSpan[] = [];
  rewriteRefs(text, (token, start) => {
    spans.push({ start, end: token.end });
    return null;
  });
  return mergeColonConnectedSpans(text, spans);
}

function readQuotedSheet(text: string, start: number): number | null {
  if (text[start] !== "'") return null;
  let i = start + 1;
  while (i < text.length) {
    if (text[i] === "'") {
      if (text[i + 1] === "'") {
        i += 2;
        continue;
      }
      return text[i + 1] === '!' ? i + 2 : null;
    }
    i += 1;
  }
  return null;
}

function r1c1ReferenceSpans(text: string): FormulaReferenceInsertionSpan[] {
  const spans: FormulaReferenceInsertionSpan[] = [];
  let i = text.startsWith('=') ? 1 : 0;
  while (i < text.length) {
    if (text[i] === '"') {
      i += 1;
      while (i < text.length) {
        if (text[i] === '"' && text[i + 1] === '"') {
          i += 2;
          continue;
        }
        if (text[i] === '"') {
          i += 1;
          break;
        }
        i += 1;
      }
      continue;
    }
    if (text[i] === '[') {
      // R1C1 relative coordinates use brackets as part of the reference atom.
      const before = text.slice(0, i);
      if (/(?:^|[^\p{L}\p{N}_.])R(?:\[-?\d+\])?C?(?:\[-?\d*)?$/iu.test(before)) {
        i += 1;
        while (i < text.length && text[i] !== ']') i += 1;
        if (text[i] === ']') i += 1;
        continue;
      }
      let depth = 1;
      i += 1;
      while (i < text.length && depth > 0) {
        if (text[i] === '[') depth += 1;
        else if (text[i] === ']') depth -= 1;
        else if (text[i] === "'" && ['[', ']', '@', '#', "'"].includes(text[i + 1] ?? '')) {
          i += 2;
          continue;
        }
        i += 1;
      }
      continue;
    }
    if (text[i] === "'") {
      const after = readQuotedSheet(text, i);
      if (after === null) {
        i += 1;
        continue;
      }
      const token = matchR1C1At(text, after);
      if (token) {
        spans.push({ start: i, end: token.end });
        i = token.end;
      } else i = after;
      continue;
    }
    if (i > 0 && IDENT_CHAR.test(text[i - 1] ?? '')) {
      i += 1;
      continue;
    }
    const unquoted = /^([\p{L}_][\p{L}\p{N}_.]*)!/u.exec(text.slice(i));
    const qualifierEnd = unquoted ? i + unquoted[0].length : i;
    const token = matchR1C1At(text, qualifierEnd);
    if (token) {
      spans.push({ start: i, end: token.end });
      i = token.end;
      continue;
    }
    i += 1;
  }
  return spans;
}

function matchR1C1At(text: string, start: number): { end: number } | null {
  R1C1_ATOM.lastIndex = start;
  const first = R1C1_ATOM.exec(text);
  if (!first || first.index !== start) return null;
  let end = start + first[0].length;
  while (true) {
    const colon = /^\s*:\s*/u.exec(text.slice(end));
    if (!colon) break;
    const nextStart = end + colon[0].length;
    R1C1_ATOM.lastIndex = nextStart;
    const next = R1C1_ATOM.exec(text);
    if (!next || next.index !== nextStart) break;
    end = nextStart + next[0].length;
  }
  if (IDENT_CHAR.test(text[end] ?? '') || text[end] === '[' || text[end] === '(') return null;
  return { end };
}

function completeReferenceAt(
  text: string,
  caret: number,
  r1c1: boolean,
): FormulaReferenceInsertionSpan | null {
  const spans = r1c1 ? r1c1ReferenceSpans(text) : a1ReferenceSpans(text);
  return spans.find((span) => caret >= span.start && caret <= span.end) ?? null;
}

function partialA1ReferenceStart(before: string, after: string): number | null {
  const match = /\$?[A-Za-z]+\$?(?:\d+)?(?::\$?[A-Za-z]+\$?(?:\d*)?)?$/.exec(before);
  if (!match) return null;
  const start = before.length - match[0].length;
  const preceding = before[start - 1] ?? '';
  if (preceding && (IDENT_CHAR.test(preceding) || preceding === "'")) return null;
  if (IDENT_CHAR.test(after[0] ?? '') || after[0] === '(' || after[0] === '[') return null;
  if (before.slice(start).endsWith('!')) return null;
  const parts = match[0].split(':');
  if (
    parts.some((atom) => {
      const letters = /^\$?([A-Za-z]+)/.exec(atom)?.[1];
      if (letters !== undefined) {
        const col = colFromLetters(letters);
        if (col < 0 || col > MAX_COL) return true;
      }
      const digits = /^\$?[A-Za-z]+\$?(\d+)/.exec(atom)?.[1];
      if (digits !== undefined) {
        const row = Number.parseInt(digits, 10);
        if (row < 1 || row > MAX_ROW + 1) return true;
      }
      return false;
    })
  )
    return null;
  return start;
}

const OPERAND_BOUNDARY = new Set([
  '=',
  '+',
  '-',
  '*',
  '/',
  '^',
  '&',
  '<',
  '>',
  ',',
  ';',
  ':',
  '(',
  '@',
]);
const NEXT_OPERAND_BOUNDARY = new Set([
  ')',
  '+',
  '-',
  '*',
  '/',
  '^',
  '&',
  '<',
  '>',
  '=',
  ',',
  ';',
  ':',
]);

function previousSignificantIndex(text: string, from: number): number {
  let index = from;
  while (index > 0 && /\s/u.test(text[index - 1] ?? '')) index -= 1;
  return index - 1;
}

function nextSignificantIndex(text: string, from: number): number {
  let index = from;
  while (index < text.length && /\s/u.test(text[index] ?? '')) index += 1;
  return index;
}

function hasValidSheetQualifierAtCaret(text: string, formulaStart: number, caret: number): boolean {
  const before = text.slice(formulaStart, caret);
  if (/(?:^|[=,+\-*/^&<>,;:(@])\s*[\p{L}_][\p{L}\p{N}_.]*!$/u.test(before)) return true;

  for (let start = formulaStart; start < caret; start += 1) {
    if (text[start] !== "'") continue;
    if (readQuotedSheet(text, start) !== caret) continue;
    const previous = previousSignificantIndex(text, start);
    const previousChar = previous >= 0 ? (text[previous] ?? '') : '';
    if (!previousChar || OPERAND_BOUNDARY.has(previousChar)) return true;
  }
  return false;
}

function isEmptyOperandSlot(text: string, start: number, end: number): boolean {
  const previous = previousSignificantIndex(text, start);
  const previousChar = previous >= 0 ? (text[previous] ?? '') : '';
  const formulaStart = text.startsWith('=') ? 1 : 0;
  if (
    !OPERAND_BOUNDARY.has(previousChar) &&
    !hasValidSheetQualifierAtCaret(text, formulaStart, start)
  )
    return false;
  const next = nextSignificantIndex(text, end);
  if (next === text.length) return true;
  return NEXT_OPERAND_BOUNDARY.has(text[next] ?? '');
}

function isExpressionSelection(text: string, start: number, end: number): boolean {
  for (let i = start; i <= end; i += 1) {
    const context = formulaCaretContext(text, i);
    if (context === 'string' || (context === 'bracket' && (i === start || i === end))) return false;
  }
  return true;
}

/** Resolve the exact span a reference insertion may replace, or return null
 *  when a collapsed caret is not at a reference or empty operand slot. */
export function resolveFormulaReferenceInsertionSpan(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  options: FormulaReferenceInputOptions = {},
): FormulaReferenceInsertionSpan | null {
  const formulaStart = text.startsWith('=') ? 1 : 0;
  const rawStart = Math.max(0, Math.min(selectionStart, selectionEnd));
  const rawEnd = Math.min(text.length, Math.max(selectionStart, selectionEnd));
  const start = Math.max(formulaStart, rawStart);
  const end = Math.max(start, rawEnd);
  if (rawStart !== rawEnd) {
    if (options.requireExpressionSelection && !isExpressionSelection(text, start, end)) return null;
    return { start, end };
  }
  if (formulaCaretContext(text, start) !== 'expression') return null;

  const complete = completeReferenceAt(text, start, options.r1c1 === true);
  if (complete) return complete;
  if (!options.r1c1) {
    const partial = partialA1ReferenceStart(text.slice(0, start), text.slice(start));
    if (partial !== null) return { start: partial, end: start };
  }
  if (IDENT_CHAR.test(text[start - 1] ?? '') || IDENT_CHAR.test(text[start] ?? '')) return null;
  if (!isEmptyOperandSlot(text, start, start)) return null;
  return { start, end: start };
}

/** Replace formula text with a picked cell/range reference. A collapsed
 * caret inside an existing ref replaces that whole token, making repeated
 * pointer-drag updates stable. Explicit selections replace exactly the
 * selection, and a caret inside literal/table text is left untouched. A
 * collapsed caret directly after a completed operand (`)`, `%`, string,
 * number) is left untouched as well. */
export function insertFormulaReference(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  ref: string,
  options: FormulaReferenceInputOptions = {},
): { text: string; caret: number } {
  const rawStart = Math.max(0, Math.min(selectionStart, selectionEnd));
  const span = resolveFormulaReferenceInsertionSpan(text, selectionStart, selectionEnd, options);
  if (!span) return { text, caret: rawStart };
  const { start, end } = span;
  const result = `${text.slice(0, start)}${ref}${text.slice(end)}`;
  return { text: result, caret: start + ref.length };
}
