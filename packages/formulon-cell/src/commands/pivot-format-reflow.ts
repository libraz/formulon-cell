import { addrKey } from '../engine/address.js';
import { PIVOT_CELL_KIND } from '../engine/pivot-cell-kind.js';
import {
  lastPivotStyleLayer,
  type PivotLayoutCell,
  type PivotReportCells,
  pivotReportTop,
  pivotStyleLayer,
} from '../engine/pivot-style.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore } from '../store/store.js';

/**
 * Direct formats on a PivotTable belong to what the cell shows, not to the
 * cell: a label to its item path, a value to its row items, column items and
 * data field. When a report's layout changes, each formatted cell becomes a
 * rule keyed that way, the report's old cells are cleared, and every new cell
 * takes the rules it matches. A rule ignores a field the new layout no longer
 * has and matches every item of a field it never named, so a format set
 * before a column field was added spreads across that field's items, and
 * comes back to one cell when the field goes. The desktop's "preserve cell
 * formatting on update" behaves this way; the engine does not expose that
 * flag, so it is taken as on.
 */

type Role = 'caption' | 'label' | 'data';

/** What a report cell shows: its role and the item each dimension holds. */
interface Identity {
  role: Role;
  /** Dimension (`r:<field>`, `c:<field>`, `d`, `gtr`, `gtc`, `h`) to its value. */
  tokens: Map<string, string>;
}

interface Rule extends Identity {
  format: CellFormat;
}

const isLabelKind = (kind: number): boolean =>
  kind === PIVOT_CELL_KIND.RowLabel || kind === PIVOT_CELL_KIND.RowSubtotal;
const isColumnKind = (kind: number): boolean =>
  kind === PIVOT_CELL_KIND.ColLabel || kind === PIVOT_CELL_KIND.ColSubtotal;

/** The identity of every position of a report below its filter block, keyed `row:col`. */
function reportIdentities(report: PivotReportCells): Map<string, Identity> {
  const { range } = report;
  const top = pivotReportTop(range, report.cells);
  const cells = report.cells.filter((c) => c.row >= top);
  const dataFields = new Set(
    cells.filter((c) => c.kind === PIVOT_CELL_KIND.Data).map((c) => c.fieldName ?? ''),
  );
  const isData = (c: PivotLayoutCell): boolean => dataFields.has(c.fieldName ?? '');
  const byRow = new Map<number, PivotLayoutCell[]>();
  for (const c of cells) byRow.set(c.row, [...(byRow.get(c.row) ?? []), c]);
  let bodyTop = range.r1 + 1;
  for (let row = top; row <= range.r1; row += 1) {
    const lead = (byRow.get(row) ?? []).reduce<PivotLayoutCell | null>(
      (a, c) => (!a || c.col < a.col ? c : a),
      null,
    );
    if (lead && lead.kind !== PIVOT_CELL_KIND.Header && !isColumnKind(lead.kind)) {
      bodyTop = row;
      break;
    }
  }
  const fieldOfColumn = new Map<number, string>();
  for (const c of cells) {
    if (c.row >= bodyTop && isData(c)) fieldOfColumn.set(c.col, c.fieldName ?? '');
  }

  const rowPath = new Map<number, Map<string, string>>();
  let stack: [string, string][] = [];
  for (let row = bodyTop; row <= range.r1; row += 1) {
    const rowCells = (byRow.get(row) ?? []).filter((c) => !isData(c));
    if (rowCells.some((c) => c.kind === PIVOT_CELL_KIND.GrandTotal)) {
      rowPath.set(row, new Map([['gtr', '']]));
      continue;
    }
    for (const c of rowCells.filter((x) => isLabelKind(x.kind)).sort((a, b) => a.col - b.col)) {
      stack = [...stack.slice(0, c.depth), [`r:${c.fieldName ?? ''}`, c.text ?? '']];
    }
    rowPath.set(row, new Map(stack));
  }

  const colPath = new Map<number, Map<string, string>>();
  let colStack: [string, string][] = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    let grand = false;
    for (let row = top; row < bodyTop; row += 1) {
      const c = (byRow.get(row) ?? []).find((x) => x.col === col);
      if (!c || isData(c)) continue;
      if (c.kind === PIVOT_CELL_KIND.GrandTotal) grand = true;
      else if (isColumnKind(c.kind)) {
        colStack = [...colStack.slice(0, c.depth), [`c:${c.fieldName ?? ''}`, c.text ?? '']];
      }
    }
    colPath.set(col, grand ? new Map([['gtc', '']]) : new Map(colStack));
  }

  const out = new Map<string, Identity>();
  for (const c of cells) {
    if (c.row < bodyTop) {
      out.set(`${c.row}:${c.col}`, { role: 'caption', tokens: new Map([['h', c.text ?? '']]) });
    } else if (!isData(c)) {
      out.set(`${c.row}:${c.col}`, { role: 'label', tokens: rowPath.get(c.row) ?? new Map() });
    }
  }
  for (let row = bodyTop; row <= range.r1; row += 1) {
    for (const [col, field] of fieldOfColumn) {
      const tokens = new Map([
        ...(rowPath.get(row) ?? []),
        ...(colPath.get(col) ?? []),
        ['d', field] as [string, string],
      ]);
      out.set(`${row}:${col}`, { role: 'data', tokens });
    }
  }
  return out;
}

/** How closely `rule` fits `cell`; null when it does not apply. A dimension
 *  the report no longer has (`present`) does not constrain the rule. */
function ruleFit(rule: Identity, cell: Identity, present: ReadonlySet<string>): number | null {
  if (rule.role !== cell.role) return null;
  let score = 0;
  for (const [dim, value] of rule.tokens) {
    const own = cell.tokens.get(dim);
    if (own === undefined) {
      if (present.has(dim)) return null;
      continue;
    }
    if (own !== value) return null;
    score += 1;
  }
  return score;
}

const position = (key: string): { row: number; col: number } => {
  const [row, col] = key.split(':').map(Number);
  return { row: row ?? 0, col: col ?? 0 };
};

/** Move the direct formats of the active sheet's reports from the layout last
 *  painted to the current one. Call after the report edit, once the store's
 *  cells hold the new layout. */
export function reflowPivotFormats(store: SpreadsheetStore, wb: WorkbookHandle): void {
  const state = store.getState();
  const sheet = state.data.sheetIndex;
  const before = lastPivotStyleLayer(wb, sheet);
  const after = pivotStyleLayer(wb, state);
  if (!before || before === after) return;
  const formats = new Map(state.format.formats);
  let changed = false;
  for (const next of after.reports) {
    const previous = before.reports.find((r) => r.pivotIndex === next.pivotIndex);
    if (!previous) continue;
    const rules: Rule[] = [];
    for (const [at, identity] of reportIdentities(previous)) {
      const key = addrKey({ sheet, ...position(at) });
      const format = formats.get(key);
      if (!format || Object.keys(format).length === 0) continue;
      rules.push({ ...identity, format });
      formats.delete(key);
      changed = true;
    }
    if (rules.length === 0) continue;
    const identities = reportIdentities(next);
    const present = new Set<string>();
    for (const identity of identities.values()) {
      for (const dim of identity.tokens.keys()) present.add(dim);
    }
    for (const [at, identity] of identities) {
      const matched = rules
        .map((rule) => ({ rule, score: ruleFit(rule, identity, present) }))
        .filter((m): m is { rule: Rule; score: number } => m.score !== null)
        .sort((a, b) => a.score - b.score);
      if (matched.length === 0) continue;
      const key = addrKey({ sheet, ...position(at) });
      let format: CellFormat = { ...formats.get(key) };
      for (const { rule } of matched) format = { ...format, ...rule.format };
      formats.set(key, format);
    }
  }
  if (changed) store.setState((s) => ({ ...s, format: { ...s.format, formats } }));
}
