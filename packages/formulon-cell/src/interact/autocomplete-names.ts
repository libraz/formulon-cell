/** Subset of `WorkbookHandle.getTables()` consumed by the structured-ref
 *  suggestion path. Only `name` and `columns` are needed. */
export interface AutocompleteTable {
  name: string;
  columns: string[];
}

/** A defined or table name offered next to function names. `detail` is the
 *  secondary line (the name's refers-to text); `table` marks a table name. */
export interface AutocompleteName {
  name: string;
  detail: string;
  table?: boolean;
}

type AutocompleteDefinedName = {
  name: string;
  formula: string;
  localSheetId: number;
  hidden?: boolean;
};

/** The parameter names of a `LAMBDA(…)` formula, or null when the formula is not a LAMBDA. */
export function lambdaParameters(formula: string): string[] | null {
  const head = /^\s*=?\s*LAMBDA\s*\(/iu.exec(formula);
  if (!head) return null;
  const args: string[] = [];
  let depth = 0;
  let start = head[0].length;
  let quoted = false;
  for (let i = start; i < formula.length; i += 1) {
    const ch = formula[i];
    if (quoted) {
      // A doubled quote is a literal quote inside the string.
      if (ch === '"' && formula[i + 1] === '"') i += 1;
      else if (ch === '"') quoted = false;
      continue;
    }
    if (ch === '"') quoted = true;
    else if (ch === '(' || ch === '{' || ch === '[') depth += 1;
    else if (depth > 0 && (ch === ')' || ch === '}' || ch === ']')) depth -= 1;
    else if (depth === 0 && (ch === ',' || ch === ';')) {
      args.push(formula.slice(start, i).trim());
      start = i + 1;
    } else if (depth === 0 && ch === ')') {
      // The last argument is the body; the ones before it are the parameters.
      return args;
    }
  }
  return null;
}

const visibleFrom = (dn: AutocompleteDefinedName, sheetIndex: number): boolean =>
  dn.name.trim() !== '' && !dn.hidden && (dn.localSheetId === -1 || dn.localSheetId === sheetIndex);

/** Names usable from `sheetIndex`: workbook-scoped and that sheet's own names,
 *  plus table names. Names scoped to another sheet are excluded, and so are
 *  LAMBDA names, which complete as functions (`collectLambdaFunctions`). */
export function collectAutocompleteNames(
  definedNames: Iterable<AutocompleteDefinedName>,
  tables: readonly AutocompleteTable[],
  sheetIndex: number,
): AutocompleteName[] {
  const out: AutocompleteName[] = [];
  for (const dn of definedNames) {
    if (!visibleFrom(dn, sheetIndex) || lambdaParameters(dn.formula)) continue;
    out.push({ name: dn.name, detail: dn.formula.startsWith('=') ? dn.formula : `=${dn.formula}` });
  }
  for (const t of tables) out.push({ name: t.name, detail: t.columns.join(', '), table: true });
  return out;
}

/** LAMBDA names usable from `sheetIndex` with their parameter names, keyed by the
 *  lower-cased name. They complete, insert and show arguments as functions do. */
export function collectLambdaFunctions(
  definedNames: Iterable<AutocompleteDefinedName>,
  sheetIndex: number,
): Map<string, { name: string; params: string[] }> {
  const out = new Map<string, { name: string; params: string[] }>();
  for (const dn of definedNames) {
    if (!visibleFrom(dn, sheetIndex)) continue;
    const params = lambdaParameters(dn.formula);
    if (params) out.set(dn.name.toLocaleLowerCase(), { name: dn.name, params });
  }
  return out;
}
