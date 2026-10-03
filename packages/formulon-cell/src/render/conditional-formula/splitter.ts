function stripOuterParens(body: string): string {
  let out = body;
  for (;;) {
    if (!out.startsWith('(') || !out.endsWith(')')) return out;
    const inner = out.slice(1, -1);
    if (splitFormulaArgs(inner) === null) return out;
    out = inner.trim();
  }
}

function splitFormulaArgs(raw: string): string[] | null {
  const args = splitFormulaArgsAllowEmpty(raw);
  return args?.every((arg) => arg.length > 0) ? args : null;
}

function splitFormulaArgsAllowEmpty(raw: string): string[] | null {
  const args: string[] = [];
  const stack: Array<')' | '}' | ']'> = [];
  let quote: '"' | "'" | null = null;
  let start = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const ch = raw[i];
    if (quote) {
      if (ch === quote) {
        if (raw[i + 1] === quote) {
          i += 1;
          continue;
        }
        quote = null;
      }
      continue;
    }

    const top = stack[stack.length - 1];
    if (top === ']') {
      if (ch === '[') {
        stack.push(']');
        continue;
      }
      if (ch === ']') {
        stack.pop();
        continue;
      }
      if (ch === "'" && "[]#'@".includes(raw[i + 1] ?? '')) {
        i += 1;
      }
      continue;
    }

    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') {
      stack.push(')');
      continue;
    }
    if (ch === '{') {
      stack.push('}');
      continue;
    }
    if (ch === '[') {
      stack.push(']');
      continue;
    }
    if (ch === ')' || ch === '}' || ch === ']') {
      if (top !== ch) return null;
      stack.pop();
      continue;
    }
    if (ch === ',' && stack.length === 0) {
      args.push(raw.slice(start, i).trim());
      start = i + 1;
    }
  }
  if (quote || stack.length !== 0) return null;
  args.push(raw.slice(start).trim());
  return args;
}

type FormulaArithmeticOp = '+' | '-' | '*' | '/' | '^' | '&';

function splitFormulaArithmetic(body: string): {
  left: string;
  op: FormulaArithmeticOp;
  right: string;
} | null {
  const operatorsByPrecedence: FormulaArithmeticOp[][] = [['&'], ['+', '-'], ['*', '/'], ['^']];
  for (const ops of operatorsByPrecedence) {
    let depth = 0;
    let quote: '"' | "'" | null = null;
    const start = ops.includes('^') ? 0 : body.length - 1;
    const end = ops.includes('^') ? body.length : -1;
    const step = ops.includes('^') ? 1 : -1;
    for (let i = start; i !== end; i += step) {
      const ch = body[i];
      if (quote) {
        if (ch === quote) quote = null;
        continue;
      }
      if (ch === '"' || ch === "'") {
        quote = ch;
        continue;
      }
      if (ch === ')') {
        depth += step < 0 ? 1 : -1;
        if (depth < 0) return null;
        continue;
      }
      if (ch === '(') {
        depth += step < 0 ? -1 : 1;
        if (depth < 0) return null;
        continue;
      }
      if (depth !== 0 || !ops.includes(ch as FormulaArithmeticOp)) continue;
      if ((ch === '+' || ch === '-') && isUnaryArithmeticSign(body, i)) continue;
      const left = body.slice(0, i).trim();
      const right = body.slice(i + 1).trim();
      if (left.length === 0 || right.length === 0) continue;
      return { left, op: ch as FormulaArithmeticOp, right };
    }
  }
  return null;
}

function isUnaryArithmeticSign(body: string, index: number): boolean {
  for (let i = index - 1; i >= 0; i -= 1) {
    const ch = body[i];
    if (ch === ' ') continue;
    return ch === '(' || ch === '+' || ch === '-' || ch === '*' || ch === '/';
  }
  return true;
}

function splitFormulaComparison(
  body: string,
): { left: string; op: '>' | '<' | '>=' | '<=' | '=' | '<>'; right: string } | null {
  let depth = 0;
  let quote: '"' | "'" | null = null;
  for (let i = 0; i < body.length; i += 1) {
    const ch = body[i];
    if (quote) {
      if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') {
      depth += 1;
      continue;
    }
    if (ch === ')') {
      depth -= 1;
      if (depth < 0) return null;
      continue;
    }
    if (depth !== 0) continue;
    const two = body.slice(i, i + 2);
    const op =
      two === '>=' || two === '<=' || two === '<>'
        ? two
        : ch === '>' || ch === '<' || ch === '='
          ? ch
          : null;
    if (!op) continue;
    const left = body.slice(0, i).trim();
    const right = body.slice(i + op.length).trim();
    return left && right ? { left, op, right } : null;
  }
  return null;
}

export {
  splitFormulaArgs,
  splitFormulaArgsAllowEmpty,
  splitFormulaArithmetic,
  splitFormulaComparison,
  stripOuterParens,
};
