/**
 * In-memory fallback "engine" that satisfies the subset of FormulonModule /
 * Workbook surface used by WorkbookHandle. Activated when the WASM load
 * fails — typically because the host page is not crossOriginIsolated and
 * the pthreaded WASM cannot allocate SharedArrayBuffer.
 *
 * It stores values literally and evaluates a tiny spreadsheet-formula subset:
 *   numbers, +, -, *, /, parens, references (A1), ranges (A1:B3),
 *   SUM, AVERAGE, MIN, MAX, COUNT, IF, AND, OR, NOT.
 *
 * Anything outside that surface returns `#NEEDS_ENGINE!`. This is enough
 * for demo / SSR / non-isolated playgrounds; production sites should
 * configure COOP+COEP and load the real WASM.
 */

import type {
  CellEntry,
  CellResult,
  EvalResult,
  FormulonModule,
  SaveResult,
  Status,
  StringResult,
  Value,
  Workbook,
} from './types.js';

const ok: Status = { ok: true, status: 0, message: '', context: '' };
const err = (m: string): Status => ({ ok: false, status: 1, message: m, context: '' });

const blankValue = (): Value => ({ kind: 0, number: 0, boolean: 0, text: '', errorCode: 0 });
const numberValue = (n: number): Value => ({
  kind: 1,
  number: n,
  boolean: 0,
  text: '',
  errorCode: 0,
});
const boolValue = (b: boolean): Value => ({
  kind: 2,
  number: 0,
  boolean: b ? 1 : 0,
  text: '',
  errorCode: 0,
});
const textValue = (s: string): Value => ({ kind: 3, number: 0, boolean: 0, text: s, errorCode: 0 });
const errorValue = (code: number): Value => ({
  kind: 4,
  number: 0,
  boolean: 0,
  text: '',
  errorCode: code,
});

interface CellStore {
  literal?: Value;
  formula?: string;
  cached?: Value;
}

const NEEDS_ENGINE = errorValue(99); // local sentinel — UI will render "#ERR!"

/**
 * Intentionally NOT `implements Workbook` — the stub only satisfies the
 * subset of the surface that `WorkbookHandle` exercises today. The
 * capability probe relies on `typeof wb.<method> === 'function'`, so
 * adding empty stubs here would lie to the probe and flip capability
 * flags on. Keeping the surface narrow makes the probe return `false`
 * for every optional capability under the stub, which is exactly the
 * behaviour the UI expects when running without the real engine.
 */
class StubWorkbook {
  private readonly sheets: { name: string; cells: Map<string, CellStore> }[] = [];

  private alive = true;

  constructor(initialSheets: string[] = ['Sheet1']) {
    for (const name of initialSheets) this.sheets.push({ name, cells: new Map() });
  }

  isValid(): boolean {
    return this.alive;
  }

  delete(): void {
    this.alive = false;
    this.sheets.length = 0;
  }

  save(): SaveResult {
    return { status: err('save unavailable in stub'), bytes: null };
  }

  addSheet(name: string): Status {
    this.sheets.push({ name, cells: new Map() });
    return ok;
  }

  sheetCount(): number {
    return this.sheets.length;
  }

  sheetName(idx: number): StringResult {
    const s = this.sheets[idx];
    return s ? { status: ok, value: s.name } : { status: err('out of range'), value: '' };
  }

  setNumber(sheet: number, row: number, col: number, value: number): Status {
    return this.put(sheet, row, col, { literal: numberValue(value) });
  }

  setBool(sheet: number, row: number, col: number, value: boolean): Status {
    return this.put(sheet, row, col, { literal: boolValue(value) });
  }

  setText(sheet: number, row: number, col: number, text: string): Status {
    return this.put(sheet, row, col, { literal: textValue(text) });
  }

  setBlank(sheet: number, row: number, col: number): Status {
    const s = this.sheets[sheet];
    if (!s) return err('sheet');
    s.cells.delete(`${row}:${col}`);
    return ok;
  }

  setFormula(sheet: number, row: number, col: number, formula: string): Status {
    return this.put(sheet, row, col, { formula });
  }

  getValue(sheet: number, row: number, col: number): CellResult {
    const s = this.sheets[sheet];
    if (!s) return { status: err('sheet'), value: blankValue() };
    const cell = s.cells.get(`${row}:${col}`);
    if (!cell) return { status: ok, value: blankValue() };
    if (cell.cached) return { status: ok, value: cell.cached };
    if (cell.literal) return { status: ok, value: cell.literal };
    return { status: ok, value: blankValue() };
  }

  recalc(): Status {
    // Iterate up to depth-of-chain times so that A→B→C reaches a stable fixed
    // point in one call. The cap is generous; real workloads shouldn't need
    // anywhere near it. We bail when nothing changed in a pass.
    const MAX_PASSES = 16;
    for (let pass = 0; pass < MAX_PASSES; pass += 1) {
      let changed = false;
      for (const s of this.sheets) {
        for (const cell of s.cells.values()) {
          if (cell.formula) {
            const prev = cell.cached;
            const next = this.evalFormula(s, cell.formula);
            cell.cached = next;
            if (!prev || prev.kind !== next.kind || JSON.stringify(prev) !== JSON.stringify(next)) {
              changed = true;
            }
          } else {
            cell.cached = cell.literal;
          }
        }
      }
      if (!changed) break;
    }
    return ok;
  }

  setIterative(): Status {
    return ok;
  }

  cellCount(sheet: number): number {
    return this.sheets[sheet]?.cells.size ?? 0;
  }

  cellAt(sheet: number, idx: number): CellEntry {
    const s = this.sheets[sheet];
    if (!s) return { status: err('sheet'), row: 0, col: 0, formula: null, value: blankValue() };
    let i = 0;
    for (const [key, cell] of s.cells) {
      if (i === idx) {
        const [r, c] = key.split(':').map(Number);
        return {
          status: ok,
          row: r ?? 0,
          col: c ?? 0,
          formula: cell.formula ?? null,
          value: cell.cached ?? cell.literal ?? blankValue(),
        };
      }
      i += 1;
    }
    return { status: err('idx'), row: 0, col: 0, formula: null, value: blankValue() };
  }

  definedNameCount(): number {
    return 0;
  }

  definedNameAt(): never {
    throw new Error('not impl');
  }

  tableCount(): number {
    return 0;
  }

  tableAt(): never {
    throw new Error('not impl');
  }

  passthroughCount(): number {
    return 0;
  }

  passthroughAt(): never {
    throw new Error('not impl');
  }

  // ---- internals -----------------------------------------------------

  private put(sheet: number, row: number, col: number, store: CellStore): Status {
    const s = this.sheets[sheet];
    if (!s) return err('sheet');
    s.cells.set(`${row}:${col}`, store);
    return ok;
  }

  private evalFormula(sheet: { cells: Map<string, CellStore>; name: string }, src: string): Value {
    try {
      const expr = src.trim().replace(/^=/, '');
      const v = parseAndEval(expr, (row, col) => {
        const cell = sheet.cells.get(`${row}:${col}`);
        if (!cell) return 0;
        const cv = cell.cached ?? cell.literal;
        if (!cv) return 0;
        if (cv.kind === 1) return cv.number;
        if (cv.kind === 2) return cv.boolean;
        if (cv.kind === 3) {
          const n = Number(cv.text);
          return Number.isFinite(n) ? n : 0;
        }
        return 0;
      });
      if (typeof v === 'number') return Number.isFinite(v) ? numberValue(v) : NEEDS_ENGINE;
      if (typeof v === 'boolean') return boolValue(v);
      if (typeof v === 'string') return textValue(v);
      return NEEDS_ENGINE;
    } catch {
      return NEEDS_ENGINE;
    }
  }
}

// --- mini parser/evaluator ------------------------------------------------

type RefResolver = (row: number, col: number) => number;
type FnArg = number | boolean | number[];

function parseAndEval(src: string, ref: RefResolver): number | boolean | string {
  let pos = 0;

  const peek = (): string => src[pos] ?? '';
  const peekAt = (n: number): string => src[pos + n] ?? '';
  const skip = (): void => {
    while (pos < src.length && /\s/.test(peek())) pos += 1;
  };

  function parseRef(): { row: number; col: number } {
    skip();
    let col = 0;
    let cur = peek().toUpperCase();
    while (cur >= 'A' && cur <= 'Z') {
      col = col * 26 + (cur.charCodeAt(0) - 64);
      pos += 1;
      cur = peek().toUpperCase();
    }
    let row = 0;
    while (peek() >= '0' && peek() <= '9') {
      row = row * 10 + Number(peek());
      pos += 1;
    }
    return { row: row - 1, col: col - 1 };
  }

  function parseRange(): number[] {
    const a = parseRef();
    if (peek() === ':') {
      pos += 1;
      const b = parseRef();
      const out: number[] = [];
      const r0 = Math.min(a.row, b.row);
      const r1 = Math.max(a.row, b.row);
      const c0 = Math.min(a.col, b.col);
      const c1 = Math.max(a.col, b.col);
      for (let r = r0; r <= r1; r += 1) for (let c = c0; c <= c1; c += 1) out.push(ref(r, c));
      return out;
    }
    return [ref(a.row, a.col)];
  }

  function parseFn(name: string): number | boolean {
    skip();
    if (peek() !== '(') throw new Error('(');
    pos += 1;
    const args: FnArg[] = [];
    skip();
    if (peek() !== ')') {
      while (true) {
        const start = pos;
        const looksRef =
          /[A-Z]/i.test(peek()) && (/[A-Z0-9]/i.test(peekAt(1)) || /[0-9]/.test(peekAt(1)));
        if (looksRef) args.push(parseRange());
        else {
          pos = start;
          args.push(parseExpr());
        }
        skip();
        if (peek() === ',') {
          pos += 1;
          skip();
          continue;
        }
        break;
      }
    }
    if (peek() !== ')') throw new Error(')');
    pos += 1;
    const flat = args.flatMap((a) => (Array.isArray(a) ? a : [a]));
    switch (name) {
      case 'SUM':
        return flat.reduce((s: number, v) => s + Number(v), 0);
      case 'AVERAGE':
      case 'AVG':
        return flat.length === 0
          ? 0
          : flat.reduce((s: number, v) => s + Number(v), 0) / flat.length;
      case 'MIN':
        return Math.min(...flat.map(Number));
      case 'MAX':
        return Math.max(...flat.map(Number));
      case 'COUNT':
        return flat.length;
      case 'IF': {
        const cond = Boolean(args[0]);
        return Number(cond ? args[1] : args[2]);
      }
      case 'AND':
        return flat.every((v) => Number(v) !== 0);
      case 'OR':
        return flat.some((v) => Number(v) !== 0);
      case 'NOT':
        return Number(args[0]) === 0;
      default:
        throw new Error(`unsupported fn ${name}`);
    }
  }

  function parsePrimary(): number | boolean {
    skip();
    const ch = peek();
    if (ch === '(') {
      pos += 1;
      const v = parseExpr();
      skip();
      if (peek() !== ')') throw new Error(')');
      pos += 1;
      return v;
    }
    if (ch === '-') {
      pos += 1;
      return -Number(parsePrimary());
    }
    if (ch === '+') {
      pos += 1;
      return Number(parsePrimary());
    }
    if (/[0-9.]/.test(ch)) {
      let s = '';
      while (/[0-9.]/.test(peek())) {
        s += peek();
        pos += 1;
      }
      if (/[eE]/.test(peek())) {
        s += peek();
        pos += 1;
        if (/[+-]/.test(peek())) {
          s += peek();
          pos += 1;
        }
        while (/[0-9]/.test(peek())) {
          s += peek();
          pos += 1;
        }
      }
      return Number(s);
    }
    if (/[A-Za-z]/.test(ch)) {
      let name = '';
      while (/[A-Za-z]/.test(peek())) {
        name += peek();
        pos += 1;
      }
      const upper = name.toUpperCase();
      skip();
      if (peek() === '(') return parseFn(upper);
      // It's a reference like A1.
      pos -= name.length;
      const ranges = parseRange();
      return Number(ranges[0] ?? 0);
    }
    throw new Error(`unexpected ${ch}`);
  }

  function parseTerm(): number | boolean {
    let v = Number(parsePrimary());
    skip();
    while (peek() === '*' || peek() === '/') {
      const op = peek();
      pos += 1;
      const rhs = Number(parsePrimary());
      v = op === '*' ? v * rhs : v / rhs;
      skip();
    }
    return v;
  }

  function parseExpr(): number | boolean {
    let v = Number(parseTerm());
    skip();
    while (peek() === '+' || peek() === '-') {
      const op = peek();
      pos += 1;
      const rhs = Number(parseTerm());
      v = op === '+' ? v + rhs : v - rhs;
      skip();
    }
    return v;
  }

  return parseExpr();
}

/** Casts a partially-conforming `StubWorkbook` to the full `Workbook`
 *  surface the rest of the codebase types against. The capability probe
 *  prevents callers from invoking the missing methods at runtime. */
const asWorkbook = (wb: StubWorkbook): Workbook => wb as unknown as Workbook;

class StubModule implements FormulonModule {
  Workbook = {
    createDefault: (): Workbook => asWorkbook(new StubWorkbook(['Sheet1'])),
    createEmpty: (): Workbook => asWorkbook(new StubWorkbook([])),
    loadBytes: (): Workbook => {
      const wb = new StubWorkbook([]);
      // immediately invalidate — the stub cannot decode .xlsx
      wb.delete();
      return asWorkbook(wb);
    },
  };

  evalFormula(formula: string): EvalResult {
    return { status: err('evalFormula unavailable in stub'), value: textValue(formula) };
  }

  versionString(): string {
    return 'stub';
  }

  statusString(): string {
    return 'kStub';
  }

  lastErrorMessage(): string {
    return '';
  }

  lastErrorContext(): string {
    return '';
  }
}

export function createStubModule(): FormulonModule {
  return new StubModule();
}
