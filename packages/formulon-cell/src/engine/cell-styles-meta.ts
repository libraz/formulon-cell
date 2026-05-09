/** One named-style entry surfaced by `computeNamedCellStyles`. Mirrors the
 *  `getCellStyle` shape but flattened with `index` so callers don't need
 *  a parallel array of indices. */
export interface NamedCellStyle {
  index: number;
  name: string;
  xfId: number;
  builtinId: number;
  iLevel: number;
  customBuiltin: boolean;
}

/** Source shape consumed by `computeNamedCellStyles`. The helper only
 *  needs a count + per-index lookup, so the function is exercisable from
 *  unit tests without spinning up the real `WorkbookHandle`. */
export interface NamedCellStylesView {
  cellStyleCount(): number;
  getCellStyle(index: number): {
    name: string;
    xfId: number;
    builtinId: number;
    iLevel: number;
    hidden: boolean;
    customBuiltin: boolean;
  } | null;
}

/** Walk every entry returned by `view.cellStyleCount()` and emit one
 *  NamedCellStyle per non-hidden record. Hidden built-ins (e.g. the
 *  legacy "Comma [0]" / "Currency [0]" entries) are filtered out to
 *  match the "Cell Styles" gallery, which only surfaces user-visible
 *  styles by default. */
export function computeNamedCellStyles(view: NamedCellStylesView): NamedCellStyle[] {
  const n = view.cellStyleCount();
  const out: NamedCellStyle[] = [];
  for (let i = 0; i < n; i += 1) {
    const s = view.getCellStyle(i);
    if (!s || s.hidden) continue;
    out.push({
      index: i,
      name: s.name,
      xfId: s.xfId,
      builtinId: s.builtinId,
      iLevel: s.iLevel,
      customBuiltin: s.customBuiltin,
    });
  }
  return out;
}
