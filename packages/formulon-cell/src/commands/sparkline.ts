import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import { mutators, type Sparkline, type SpreadsheetStore } from '../store/store.js';
import { adjustInterval } from './axis-shift.js';
import { adjustFormulaForRowColEdit } from './formula-refs.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';
import { recordSparklineChange } from './slice-history.js';

export interface SparklineEntry {
  addr: Addr;
  spec: Sparkline;
}

export function listSparklines(state: {
  sparkline: { sparklines: ReadonlyMap<string, Sparkline> };
}): readonly SparklineEntry[] {
  const out: SparklineEntry[] = [];
  for (const [key, spec] of state.sparkline.sparklines) {
    const addr = parseAddrKey(key);
    if (addr) out.push({ addr, spec: { ...spec } });
  }
  return out.sort(
    (a, b) => a.addr.sheet - b.addr.sheet || a.addr.row - b.addr.row || a.addr.col - b.addr.col,
  );
}

export function sparklineAt(
  state: { sparkline: { sparklines: ReadonlyMap<string, Sparkline> } },
  addr: Addr,
): Sparkline | null {
  const spec = state.sparkline.sparklines.get(addrKey(addr));
  return spec ? { ...spec } : null;
}

export function setSparkline(
  store: SpreadsheetStore,
  addr: Addr,
  spec: Sparkline,
  history: History | null = null,
): boolean {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return false;
  }
  recordSparklineChange(history, store, () => {
    mutators.setSparkline(store, addr, spec);
  });
  return true;
}

export function clearSparkline(
  store: SpreadsheetStore,
  addr: Addr,
  history: History | null = null,
): boolean {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return false;
  }
  recordSparklineChange(history, store, () => {
    mutators.clearSparkline(store, addr);
  });
  return true;
}

export function clearSparklinesInRange(
  store: SpreadsheetStore,
  range: Range,
  history: History | null = null,
): number {
  const state = store.getState();
  const targets = listSparklines(state)
    .filter(({ addr }) => rangeContainsAddr(range, addr))
    .map(({ addr }) => addr);
  if (targets.length === 0) return 0;
  const writable = targets.filter((addr) => {
    if (isCellWritable(state, addr)) return true;
    warnProtected(addr);
    return false;
  });
  if (writable.length === 0) return 0;
  recordSparklineChange(history, store, () => {
    for (const addr of writable) mutators.clearSparkline(store, addr);
  });
  return writable.length;
}

/** Split `Sheet2!B2:B9` / `'My Sheet'!B2:B9` into its qualifier (with the `!`)
 *  and the A1 body; the qualifier is empty for a bare range. */
const splitSourceQualifier = (source: string): { qualifier: string; body: string } => {
  const bang = source.lastIndexOf('!');
  return bang < 0
    ? { qualifier: '', body: source }
    : { qualifier: source.slice(0, bang + 1), body: source.slice(bang + 1) };
};

const unquoteSheetName = (qualifier: string): string => {
  const name = qualifier.slice(0, -1);
  return name.startsWith("'") && name.endsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
};

const quoteSheetName = (name: string): string =>
  /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;

/** Re-point sparklines after a row/column insert or delete on `sheet`: host
 *  cells move with their cells (a deleted host drops its sparkline) and a
 *  source range on that sheet follows its data. `sheetName` is the edited
 *  sheet's name, used to recognise qualified sources. */
export function shiftSparklinesForAxisEdit(
  sparklines: ReadonlyMap<string, Sparkline>,
  sheet: number,
  sheetName: string,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): Map<string, Sparkline> {
  const out = new Map<string, Sparkline>();
  const target = sheetName.toLowerCase();
  for (const [key, spec] of sparklines) {
    const host = parseAddrKey(key);
    if (!host) {
      out.set(key, spec);
      continue;
    }
    const { qualifier, body } = splitSourceQualifier(spec.source);
    const sourceOnSheet = qualifier
      ? unquoteSheetName(qualifier).toLowerCase() === target
      : host.sheet === sheet;
    const source = sourceOnSheet
      ? `${qualifier}${adjustFormulaForRowColEdit(`=${body}`, axis, split, delta).slice(1)}`
      : spec.source;
    const next = source === spec.source ? spec : { ...spec, source };
    if (host.sheet !== sheet) {
      out.set(key, next);
      continue;
    }
    const at = axis === 'row' ? host.row : host.col;
    const moved = adjustInterval(at, at, split, delta);
    if (!moved) continue;
    out.set(
      addrKey({
        ...host,
        row: axis === 'row' ? moved[0] : host.row,
        col: axis === 'col' ? moved[0] : host.col,
      }),
      next,
    );
  }
  return out;
}

/** Rewrite sparkline sources qualified with a renamed sheet's old name. */
export function renameSheetInSparklineSources(
  sparklines: ReadonlyMap<string, Sparkline>,
  oldName: string,
  newName: string,
): Map<string, Sparkline> | null {
  const old = oldName.toLowerCase();
  let changed = false;
  const out = new Map<string, Sparkline>();
  for (const [key, spec] of sparklines) {
    const { qualifier, body } = splitSourceQualifier(spec.source);
    if (qualifier && unquoteSheetName(qualifier).toLowerCase() === old) {
      changed = true;
      out.set(key, { ...spec, source: `${quoteSheetName(newName)}!${body}` });
    } else {
      out.set(key, spec);
    }
  }
  return changed ? out : null;
}
