/**
 * In-memory test engine that satisfies the subset of the `Formulon` /
 * `Workbook` / `Sheet` surface used by WorkbookHandle. Activated only by
 * preferStub; WASM initialization failures are reported to the caller.
 *
 * It stores values literally and evaluates a tiny spreadsheet-formula subset:
 *   numbers, +, -, *, /, parens, references (A1), ranges (A1:B3),
 *   SUM, AVERAGE, MIN, MAX, COUNT, IF, AND, OR, NOT.
 *
 * Anything outside that surface returns `#NEEDS_ENGINE!`. This is enough
 * for explicit tests and demos; production sites should load the real WASM.
 */

import { type ErrorCode, FormulonError, StatusCode, ValueKind } from '@libraz/formulon';
import type { Formulon, RangeCell, Sheet, Value, Workbook } from './types.js';

// Engine values are read inside functions only: a test that mocks the engine
// package must still be able to load this module.
const blankValue = (): Value => ({ kind: ValueKind.Blank });
const numberValue = (n: number): Value => ({ kind: ValueKind.Number, number: n });
const boolValue = (b: boolean): Value => ({ kind: ValueKind.Bool, boolean: b });
const textValue = (s: string): Value => ({ kind: ValueKind.Text, text: s });
const errorValue = (code: number): Value => ({
  kind: ValueKind.Error,
  errorCode: code as ErrorCode,
});

/** Local sentinel outside the engine's error codes; the UI renders "#ERR!". */
const NEEDS_ENGINE_CODE = 99;

const fail = (code: StatusCode, message: string): FormulonError => new FormulonError(code, message);

/** Keys a host or a test runner probes on any object; they read as absent rather than throwing. */
const PROBE_KEYS: ReadonlySet<string> = new Set([
  'then',
  'toJSON',
  'constructor',
  'asymmetricMatch',
  'nodeType',
  'tagName',
]);

/** A stand-in for an engine member the stub lacks: calling it, or reading through it, throws
 *  `NotImplemented` naming the member. */
const unavailable = (member: string): unknown => {
  const raise = (): never => {
    throw fail(StatusCode.NotImplemented, `${member} unavailable in the stub engine`);
  };
  return new Proxy(raise, {
    get: (target, key) => (typeof key === 'symbol' ? Reflect.get(target, key) : raise()),
    apply: raise,
  });
};

/** Wraps a stub object so a member it does not implement throws `NotImplemented` at the call,
 *  as the engine would, instead of reading as `undefined`. `in` still reports the real members. */
const withUnavailableMembers = <T extends object>(target: T): T =>
  new Proxy(target, {
    get: (obj, key, receiver) => {
      if (typeof key === 'symbol' || key in obj || PROBE_KEYS.has(key) || /^[@$]{2}/.test(key)) {
        return Reflect.get(obj, key, receiver);
      }
      return unavailable(key);
    },
  });

interface CellStore {
  literal?: Value;
  formula?: string;
  cached?: Value;
}

interface CellIndexEntry {
  key: string;
  row: number;
  col: number;
}

interface SheetData {
  name: string;
  cells: Map<string, CellStore>;
  cellIndex: CellIndexEntry[] | null;
}

/** Native form of a stored value, as `RangeCell.value` carries it. */
const nativeOf = (v: Value): RangeCell['value'] => {
  switch (v.kind) {
    case ValueKind.Number:
      return v.number;
    case ValueKind.Bool:
      return v.boolean;
    case ValueKind.Text:
      return v.text;
    case ValueKind.Error:
      return { code: v.errorCode } as RangeCell['value'];
    default:
      return null;
  }
};

/**
 * Intentionally NOT `implements Sheet` — the stub only satisfies the subset
 * of the surface that `WorkbookHandle` exercises. Optional features are
 * reported absent by the stub's capability set, so callers never reach
 * the missing methods.
 */
class StubSheet {
  constructor(
    private readonly book: StubWorkbook,
    private readonly data: SheetData,
  ) {}

  get workbook(): Workbook {
    return asWorkbook(this.book);
  }

  get index(): number {
    return this.book.indexOf(this.data);
  }

  get name(): string {
    return this.live().name;
  }

  setNumber(row: number, col: number, value: number): void {
    this.put(row, col, { literal: numberValue(value) });
  }

  setBool(row: number, col: number, value: boolean): void {
    this.put(row, col, { literal: boolValue(value) });
  }

  setError(row: number, col: number, errorCode: number): void {
    this.put(row, col, { literal: errorValue(errorCode) });
  }

  setText(row: number, col: number, text: string): void {
    this.put(row, col, { literal: textValue(text) });
  }

  setBlank(row: number, col: number): void {
    const s = this.live();
    if (s.cells.delete(`${row}:${col}`)) s.cellIndex = null;
  }

  setFormula(row: number, col: number, formula: string): void {
    this.put(row, col, { formula: formula.startsWith('=') ? formula : `=${formula}` });
  }

  getTypedValue(row: number, col: number): Value {
    const cell = this.live().cells.get(`${row}:${col}`);
    return cell?.cached ?? cell?.literal ?? blankValue();
  }

  /** Every stored cell in row-major order; the range argument is not modelled. */
  listCells(): readonly RangeCell[] {
    const s = this.live();
    if (s.cellIndex === null) {
      s.cellIndex = [];
      for (const key of s.cells.keys()) {
        const [row, col] = key.split(':').map(Number);
        s.cellIndex.push({ key, row: row ?? 0, col: col ?? 0 });
      }
      s.cellIndex.sort((a, b) => a.row - b.row || a.col - b.col);
    }
    const out: RangeCell[] = [];
    for (const indexed of s.cellIndex) {
      const cell = s.cells.get(indexed.key);
      if (!cell) continue;
      const value = cell.cached ?? cell.literal ?? blankValue();
      out.push({
        row: indexed.row,
        col: indexed.col,
        formula: cell.formula ?? null,
        kind: value.kind,
        value: nativeOf(value),
      });
    }
    return out;
  }

  listTables(): readonly never[] {
    this.live();
    return [];
  }

  /** The sheet's data; throws once the sheet is gone, like an engine handle. */
  private live(): SheetData {
    this.book.indexOf(this.data);
    return this.data;
  }

  private put(row: number, col: number, store: CellStore): void {
    const s = this.live();
    const key = `${row}:${col}`;
    const existed = s.cells.has(key);
    s.cells.set(key, store);
    if (!existed) s.cellIndex = null;
  }
}

/** Intentionally NOT `implements Workbook`; see `StubSheet`. */
class StubWorkbook {
  private readonly sheets: SheetData[] = [];

  /** One handle per sheet, as the engine guarantees. */
  private readonly handles = new Map<SheetData, StubSheet>();

  private alive = true;

  readonly ooxml = {
    listPassthroughs: (): readonly string[] => [],
  };

  constructor(initialSheets: string[] = ['Sheet1']) {
    for (const name of initialSheets) {
      this.sheets.push({ name, cells: new Map(), cellIndex: null });
    }
  }

  get disposed(): boolean {
    return !this.alive;
  }

  dispose(): void {
    this.alive = false;
    this.sheets.length = 0;
    this.handles.clear();
  }

  get sheetCount(): number {
    return this.sheets.length;
  }

  hasSheet(ref: number | string): boolean {
    return this.find(ref) !== undefined;
  }

  sheet(ref: number | string): Sheet {
    const data = this.find(ref);
    if (!data) throw fail(StatusCode.NotFound, `no sheet ${String(ref)}`);
    let handle = this.handles.get(data);
    if (!handle) {
      handle = withUnavailableMembers(new StubSheet(this, data));
      this.handles.set(data, handle);
    }
    return handle as unknown as Sheet;
  }

  listSheets(): readonly Sheet[] {
    return this.sheets.map((_, i) => this.sheet(i));
  }

  /** Position of `data`; throws once its sheet is gone. */
  indexOf(data: SheetData): number {
    const i = this.sheets.indexOf(data);
    if (i < 0) throw fail(StatusCode.NotFound, 'sheet was removed');
    return i;
  }

  save(): Uint8Array {
    throw fail(StatusCode.NotImplemented, 'save unavailable in stub');
  }

  saveWithDiagnostics(): never {
    throw fail(StatusCode.NotImplemented, 'save unavailable in stub');
  }

  addSheet(name?: string | null): Sheet {
    const proposed = name ?? `Sheet${this.sheets.length + 1}`;
    if (this.find(proposed)) throw fail(StatusCode.AlreadyExists, `sheet ${proposed} exists`);
    this.sheets.push({ name: proposed, cells: new Map(), cellIndex: null });
    return this.sheet(this.sheets.length - 1);
  }

  /** Every formula is recomputed on each pass, so a full recalculation is the same pass. */
  recalcFull(): void {
    this.recalc();
  }

  recalc(): void {
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
  }

  updateIterative(): void {}

  listDefinedNames(): readonly never[] {
    return [];
  }

  // ---- internals -----------------------------------------------------

  /** By index, else by name: an exact match, else the single case-folded match (the first two
   *  tiers of the engine's name identity; the stub skips its width and kana folding). */
  private find(ref: number | string): SheetData | undefined {
    if (typeof ref === 'number') return this.sheets[ref];
    const exact = this.sheets.find((s) => s.name === ref);
    if (exact) return exact;
    const folded = ref.toLowerCase();
    const matches = this.sheets.filter((s) => s.name.toLowerCase() === folded);
    return matches.length === 1 ? matches[0] : undefined;
  }

  private evalFormula(sheet: SheetData, src: string): Value {
    try {
      const expr = src.trim().replace(/^=/, '');
      const v = parseAndEval(expr, (row, col) => {
        const cell = sheet.cells.get(`${row}:${col}`);
        if (!cell) return 0;
        const cv = cell.cached ?? cell.literal;
        if (!cv) return 0;
        if (cv.kind === ValueKind.Number) return cv.number;
        if (cv.kind === ValueKind.Bool) return cv.boolean ? 1 : 0;
        if (cv.kind === ValueKind.Text) {
          const n = Number(cv.text);
          return Number.isFinite(n) ? n : 0;
        }
        return 0;
      });
      if (typeof v === 'number')
        return Number.isFinite(v) ? numberValue(v) : errorValue(NEEDS_ENGINE_CODE);
      if (typeof v === 'boolean') return boolValue(v);
      if (typeof v === 'string') return textValue(v);
      return errorValue(NEEDS_ENGINE_CODE);
    } catch {
      return errorValue(NEEDS_ENGINE_CODE);
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
 *  surface the rest of the codebase types against. The stub's capability
 *  set keeps callers off the missing methods at runtime. */
const asWorkbook = (wb: StubWorkbook): Workbook => wb as unknown as Workbook;

const stubModule = {
  Workbook: {
    /** A stub book carries no profile defaults: the stub has no profile data, so its font,
     *  row height and column width are the stub's own for every locale and platform. */
    createDefault: (): Workbook => asWorkbook(withUnavailableMembers(new StubWorkbook(['Sheet1']))),
    createEmpty: (): Workbook => asWorkbook(withUnavailableMembers(new StubWorkbook([]))),
    createForProfile: (): Workbook => {
      // Profile-specific evaluation requires the real engine.
      throw fail(StatusCode.NotImplemented, 'createForProfile unavailable in stub');
    },
    load: (): Workbook => {
      throw fail(StatusCode.NotImplemented, 'the stub cannot decode workbook files');
    },
  },

  evaluateFormula(): never {
    throw fail(StatusCode.NotImplemented, 'evaluateFormula unavailable in stub');
  },

  getVersion(): string {
    return 'stub';
  },

  /** The stub emits no structured log records, so both knobs are accepted
   *  and ignored rather than reported as failures. */
  setLogMinLevel(): void {},

  setLogSink(): void {},
};

export function createStubModule(): Formulon {
  return { ...stubModule, Workbook: { ...stubModule.Workbook } } as unknown as Formulon;
}
