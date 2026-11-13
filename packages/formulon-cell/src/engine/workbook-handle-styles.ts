import {
  type CellXfInput,
  type Color,
  type ColorContext,
  type EffectiveStyle,
  type ErrorCode,
  FontScheme,
  type ThemeFonts,
  type ThemeScriptFont,
  type ThemeSource,
  type Value,
  ValueKind,
} from '@libraz/formulon';
import { computeNamedCellStyles, type NamedCellStyle } from './cell-styles-meta.js';
import { orFallback } from './result.js';
import { completeBorderRecord, completeFillRecord, completeFontRecord } from './style-records.js';
import type {
  Addr,
  BorderRecord,
  CellRect,
  CellValue,
  CellXf,
  EngineCapabilities,
  FillRecord,
  FontRecord,
  Range,
  Sheet,
  SheetFontAxisSnapshot,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

type NamedCellStyleRecord = {
  name: string;
  xfId: number;
  builtinId: number | null;
  iLevel?: number;
  hidden?: boolean;
  customBuiltin?: boolean;
};

/** The workbook theme as the handle reports it: colours in clrScheme order,
 *  the Latin / East Asian faces and each font set's per-script faces. */
export interface WorkbookTheme {
  source: ThemeSource;
  colors: number[];
  fonts: ThemeFonts;
  majorScripts: ThemeScriptFont[];
  minorScripts: ThemeScriptFont[];
}

/** `getCellStyle`'s sentinel for "this style has no OOXML built-in id". */
const CUSTOM_CELL_STYLE_BUILTIN_ID = 0xffffffff;

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleStylesMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

function ws(handle: unknown, sheet: number): Sheet {
  return internals(handle).wb.sheet(sheet);
}

const cellRect = (range: Range): CellRect => ({
  firstRow: range.r0,
  firstCol: range.c0,
  lastRow: range.r1,
  lastCol: range.c1,
});

/** Theme read per handle; the engine theme only changes through the setters
 *  below, and painting resolves a theme-linked font for every visible cell. */
const themeCache = new WeakMap<object, WorkbookTheme | null>();

/** ISO 15924 theme script a font `charset` selects; any other charset takes the latin face. */
const SCRIPT_FOR_CHARSET: ReadonlyMap<number, string> = new Map([
  [128, 'Jpan'],
  [129, 'Hang'],
  [134, 'Hans'],
  [136, 'Hant'],
  [222, 'Thai'],
]);

/** Theme script of the editing language a profile models; a theme-linked font
 *  without a charset takes that script's face, as a Japanese host shows a
 *  minor-scheme Calibri record in the theme's Jpan typeface. */
const SCRIPT_FOR_PROFILE_LOCALE: ReadonlyMap<string, string> = new Map([
  ['ja_JP', 'Jpan'],
  ['ko_KR', 'Hang'],
  ['zh_CN', 'Hans'],
  ['zh_TW', 'Hant'],
  ['th_TH', 'Thai'],
]);

function profileScript(handle: WorkbookHandleStylesMethods): string | undefined {
  const wbh = handle as unknown as WorkbookHandle;
  const profile = wbh.fontScriptProfile ?? wbh.spreadsheetProfileId();
  const locale = profile?.slice(profile.indexOf('-') + 1);
  return locale ? SCRIPT_FOR_PROFILE_LOCALE.get(locale) : undefined;
}

function readTheme(handle: WorkbookHandleStylesMethods): WorkbookTheme | null {
  if (themeCache.has(handle)) return themeCache.get(handle) ?? null;
  const native = wb(handle);
  const theme = orFallback(null, (): WorkbookTheme => {
    const colors = native.getThemeColors();
    return {
      source: colors.source,
      colors: [...colors.colors],
      fonts: native.getThemeFonts(),
      majorScripts: [...native.listThemeScriptFonts(FontScheme.Major)],
      minorScripts: [...native.listThemeScriptFonts(FontScheme.Minor)],
    };
  });
  themeCache.set(handle, theme);
  return theme;
}

function resolveArgb(handle: unknown, spec: Color, context: ColorContext): number | null {
  const capabilities = (handle as { capabilities?: EngineCapabilities }).capabilities;
  if (spec.kind === 0 || !capabilities?.colorResolution) return null;
  return orFallback(null, () => wb(handle).resolveColor(spec, context).argb);
}

/** A resolved colour keeps its selector and carries the display ARGB in `argb`. */
const withArgb = (color: Color, argb: number): Color => ({ ...color, argb });

const engineValueFromCell = (value: CellValue): Value => {
  switch (value.kind) {
    case 'number':
      return { kind: ValueKind.Number, number: value.value };
    case 'bool':
      return { kind: ValueKind.Bool, boolean: value.value };
    case 'text':
      return { kind: ValueKind.Text, text: value.value };
    case 'error':
      return { kind: ValueKind.Error, errorCode: value.code as ErrorCode };
    default:
      return { kind: ValueKind.Blank };
  }
};

export abstract class WorkbookHandleStylesMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Set the native cell-style XF for an inclusive column range. Width,
   * hidden and outline metadata stay owned by the engine's column record. */
  setColumnStyle(sheet: number, first: number, last: number, xfIndex: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.setColumnXfIndex(first, last, xfIndex);
      return true;
    });
  }

  /** Clear the native cell-style XF for an inclusive column range. */
  clearColumnStyle(sheet: number, first: number, last: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.resetColumnXfIndex(first, last);
      return true;
    });
  }

  /** Set the native cell-style XF for one worksheet row. */
  setRowStyle(sheet: number, row: number, xfIndex: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.setRowXfIndex(row, row, xfIndex);
      return true;
    });
  }

  /** Clear the native cell-style XF for one worksheet row. */
  clearRowStyle(sheet: number, row: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.resetRowXfIndex(row, row);
      return true;
    });
  }

  setRowStyleRange(sheet: number, first: number, last: number, xfIndex: number): boolean {
    assertAlive(this);
    if (!this.capabilities.rowStyleRanges) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.setRowXfIndex(first, last, xfIndex);
      return true;
    });
  }

  clearRowStyleRange(sheet: number, first: number, last: number): boolean {
    assertAlive(this);
    if (!this.capabilities.rowStyleRanges) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.resetRowXfIndex(first, last);
      return true;
    });
  }

  getRangeStyleBases(range: Range): readonly number[] | null {
    assertAlive(this);
    if (!this.capabilities.rangeStylePatch) return null;
    return orFallback(null, () => [
      ...ws(this, range.sheet).ooxml.getRangeXfBases(cellRect(range)),
    ]);
  }

  remapRangeStyle(range: Range, fromXfs: readonly number[], toXfs: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.rangeStylePatch) return false;
    if (fromXfs.length !== toXfs.length) return false;
    const mapping = new Map<number, number>();
    for (const [i, from] of fromXfs.entries()) mapping.set(from, toXfs[i] as number);
    // A repeated source XF is an ambiguous remap; refuse it rather than merge.
    if (mapping.size !== fromXfs.length) return false;
    return orFallback(false, () => {
      ws(this, range.sheet).ooxml.remapRangeXfs(cellRect(range), mapping);
      return true;
    });
  }

  snapshotRangeStyle(range: Range): Uint8Array | null {
    assertAlive(this);
    if (!this.capabilities.rangeStyleSnapshot) return null;
    return orFallback(
      null,
      () => new Uint8Array(ws(this, range.sheet).ooxml.snapshotRangeXfs(cellRect(range))),
    );
  }

  restoreRangeStyle(sheet: number, bytes: Uint8Array): boolean {
    assertAlive(this);
    if (!this.capabilities.rangeStyleSnapshot) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.restoreRangeXfs(bytes);
      return true;
    });
  }

  /** Capture only style-bearing axis records. Restoring these after clearing
   * an edited axis leaves width, hidden, outline, height and customHeight in
   * the native layout untouched. */
  getSheetFontAxisSnapshot(sheet: number): SheetFontAxisSnapshot | null {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return null;
    return orFallback(null, (): SheetFontAxisSnapshot => {
      const sheetHandle = ws(this, sheet);
      const columns: { first: number; last: number; styleXf: number }[] = [];
      // A column span is homogeneous, so its first column states the span's XF.
      for (const span of sheetHandle.listColumnOverrides()) {
        const styleXf = sheetHandle.ooxml.getColumnXfIndex(span.first);
        if (styleXf !== null) columns.push({ first: span.first, last: span.last, styleXf });
      }
      const rows: { row: number; styleXf: number }[] = [];
      for (const entry of sheetHandle.listRowOverrides()) {
        const styleXf = sheetHandle.ooxml.getRowXfIndex(entry.row);
        if (styleXf !== null) rows.push({ row: entry.row, styleXf });
      }
      return { columns, rows };
    });
  }

  /** Read the XF (eXtended Format) table index assigned to `(sheet, row, col)`.
   *  Returns 0 (the workbook's default XF row) on missing cells. Returns null
   *  when cell formatting is unavailable or the engine refuses the address. */
  getCellXfIndex(sheet: number, row: number, col: number): number | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    return orFallback(null, () => ws(this, sheet).ooxml.getCellXfIndex(row, col) ?? 0);
  }

  /** Every stored cell in `range` (default: the whole sheet) whose own cell XF is
   *  not 0, styled blanks included, row-major. `[]` when cell formatting is
   *  unavailable or the engine refuses. */
  cellXfIndices(sheet: number, range?: Range): readonly { addr: Addr; xfIndex: number }[] {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return [];
    return orFallback([], () =>
      ws(this, sheet)
        .ooxml.listCellXfIndices(range ? cellRect(range) : undefined)
        .map((e) => ({ addr: { sheet, row: e.row, col: e.col }, xfIndex: e.xfIndex })),
    );
  }

  /** Pin the XF index of `(sheet, row, col)` to `xfIndex`. The index must
   *  point at an existing row in the workbook's XF table; use `addXfRecord`
   *  first when a new row is needed. This is mainly useful for cloning
   *  formatting from one cell to another (Format-Painter parity at the
   *  engine layer) or for clearing back to xfIndex 0 (the default). Returns
   *  false on engine failure or when `capabilities.cellFormatting` is off. */
  setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): boolean {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return false;
    return orFallback(false, () => {
      ws(this, sheet).ooxml.setCellXfIndex(row, col, xfIndex);
      return true;
    });
  }

  /** Remove a direct XF marker. Blank cells inherit their row/column style;
   * nonblank cells retain the native Cell/Normal resolution contract.
   * Values, formulas and annotations remain intact. */
  clearCellXfIndex(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    return orFallback(false, () => {
      ws(this, sheet).ooxml.resetCellXfIndex(row, col);
      return true;
    });
  }

  /** Resolve the XF record at `xfIndex` to its component table indices
   *  (font / fill / border / number-format) plus alignment + wrap flags.
   *  Resolve those component indices with `getFontRecord`, `getFillRecord`,
   *  `getBorderRecord`, and `getNumFmtCode`. Returns null on engine failure
   *  or capability off. */
  getCellXf(xfIndex: number): CellXf | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    return orFallback(null, () => wb(this).ooxml.getCellXf(xfIndex));
  }

  /** Resolve a font index to its plain-data record. Returns null on engine
   *  failure or when `capabilities.cellFormatting` is off. */
  getFontRecord(fontIndex: number): FontRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    return orFallback(null, () => wb(this).ooxml.getFont(fontIndex));
  }

  /** Resolve a theme, indexed or automatic colour to opaque ARGB against the
   *  workbook theme and palette; null when the engine cannot resolve it. */
  resolveColorArgb(spec: Color, context: ColorContext = 0): number | null {
    assertAlive(this);
    return resolveArgb(this, spec, context);
  }

  /** Resolve a font record's theme, indexed, or automatic colour for display
   *  while retaining its raw selector: the resolved ARGB is `color.argb`.
   *  Theme-linked font names are projected from the current workbook theme
   *  for rendering; the returned `scheme` still preserves the source link so
   *  writeback can keep it linked. */
  getResolvedFontRecord(fontIndex: number): FontRecord | null {
    const record = this.getFontRecord(fontIndex);
    if (!record) return null;
    let resolved = record;
    if (record.scheme === FontScheme.Major || record.scheme === FontScheme.Minor) {
      const theme = this.capabilities.workbookTheme ? readTheme(this) : null;
      const major = record.scheme === FontScheme.Major;
      const script =
        record.charset !== null && record.charset !== undefined
          ? SCRIPT_FOR_CHARSET.get(record.charset)
          : profileScript(this);
      const scriptFace = script
        ? (major ? theme?.majorScripts : theme?.minorScripts)?.find((s) => s.script === script)
            ?.typeface
        : undefined;
      const themeName = scriptFace || (major ? theme?.fonts.majorLatin : theme?.fonts.minorLatin);
      if (themeName) resolved = { ...resolved, name: themeName };
    }
    if (!record.color) return resolved;
    const argb = resolveArgb(this, record.color, 0);
    return argb === null ? resolved : { ...resolved, color: withArgb(record.color, argb) };
  }

  /** Resolve a fill index to its plain-data record. */
  getFillRecord(fillIndex: number): FillRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    return orFallback(null, () => wb(this).ooxml.getFill(fillIndex));
  }

  /** Resolve a fill record's theme, indexed, or automatic colours for display
   *  while retaining both raw selectors (the resolved ARGB is each colour's
   *  `argb`). RGB records are returned unchanged. */
  getResolvedFillRecord(fillIndex: number): FillRecord | null {
    const record = this.getFillRecord(fillIndex);
    if (!record) return null;
    const fgArgb = record.foreground ? resolveArgb(this, record.foreground, 1) : null;
    const bgArgb = record.background ? resolveArgb(this, record.background, 2) : null;
    const gradient = record.gradient && {
      ...record.gradient,
      stops: record.gradient.stops.map((stop) => {
        const argb = stop.color ? resolveArgb(this, stop.color, 1) : null;
        return argb === null || !stop.color ? stop : { ...stop, color: withArgb(stop.color, argb) };
      }),
    };
    if (fgArgb === null && bgArgb === null && !gradient) return record;
    return {
      ...record,
      ...(fgArgb === null || !record.foreground
        ? {}
        : { foreground: withArgb(record.foreground, fgArgb) }),
      ...(bgArgb === null || !record.background
        ? {}
        : { background: withArgb(record.background, bgArgb) }),
      ...(gradient ? { gradient } : {}),
    };
  }

  /** Resolve a border index to its plain-data record. */
  getBorderRecord(borderIndex: number): BorderRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    return orFallback(null, () => wb(this).ooxml.getBorder(borderIndex));
  }

  /** Resolve a border record's theme, indexed, or automatic colours for
   *  display while retaining each side's raw selector. */
  getResolvedBorderRecord(borderIndex: number): BorderRecord | null {
    const record = this.getBorderRecord(borderIndex);
    if (!record) return null;
    let changed = false;
    const sides = (['left', 'right', 'top', 'bottom', 'diagonal'] as const).reduce(
      (out, side) => {
        const current = record[side];
        const argb = current.color ? resolveArgb(this, current.color, 3) : null;
        if (argb === null || !current.color) {
          out[side] = current;
        } else {
          changed = true;
          out[side] = { ...current, color: withArgb(current.color, argb) };
        }
        return out;
      },
      {} as Pick<BorderRecord, 'left' | 'right' | 'top' | 'bottom' | 'diagonal'>,
    );
    return changed ? { ...record, ...sides } : record;
  }

  /** Resolve a number-format id to its format-code string. */
  getNumFmtCode(numFmtId: number): string | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    return orFallback(null, () => wb(this).ooxml.getNumberFormat(numFmtId));
  }

  /** Read the workbook theme without exposing the raw WASM workbook. The
   *  returned selectors stay in clrScheme order so callers can edit a single
   *  preset and send the same twelve-entry array back to the engine. */
  getWorkbookTheme(): WorkbookTheme | null {
    assertAlive(this);
    if (!this.capabilities.workbookTheme) return null;
    const result = readTheme(this);
    if (!result) return null;
    return {
      ...result,
      colors: [...result.colors],
      fonts: { ...result.fonts },
      majorScripts: result.majorScripts.map((s) => ({ ...s })),
      minorScripts: result.minorScripts.map((s) => ({ ...s })),
    };
  }

  /** Native-name alias for consumers that work directly with Workbook APIs. */
  getTheme(): WorkbookTheme | null {
    return this.getWorkbookTheme();
  }

  /** Replace the workbook's twelve clrScheme colours. */
  setWorkbookThemeColors(colors: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.workbookTheme || colors.length !== 12) return false;
    themeCache.delete(this);
    const hex = colors.map((argb) => `#${(argb >>> 0).toString(16).padStart(8, '0')}`);
    return orFallback(false, () => {
      wb(this).setThemeColors({ colors: hex });
      return true;
    });
  }

  /** Native-name alias for workbook theme editing. */
  setThemeColors(colors: readonly number[]): boolean {
    return this.setWorkbookThemeColors(colors);
  }

  /** Replace the workbook's major/minor Latin and East Asian theme faces. */
  setWorkbookThemeFonts(fonts: ThemeFonts): boolean {
    assertAlive(this);
    if (!this.capabilities.workbookTheme) return false;
    themeCache.delete(this);
    return orFallback(false, () => {
      wb(this).setThemeFonts({ ...fonts });
      return true;
    });
  }

  /** Native-name alias for workbook theme editing. */
  setThemeFonts(fonts: ThemeFonts): boolean {
    return this.setWorkbookThemeFonts(fonts);
  }

  /** Return the XF selected by cell, row, column, or workbook default. This
   *  remains a read-only projection: callers must not feed its XF index back
   *  to `setCellXfIndex` just because a row/column supplied the formatting. */
  getEffectiveStyle(sheet: number, row: number, col: number): EffectiveStyle | null {
    assertAlive(this);
    if (!this.capabilities.effectiveStyle) return null;
    return orFallback(null, () => ws(this, sheet).ooxml.getEffectiveStyle(row, col));
  }

  /** Use formulon's workbook-aware display renderer for an existing cell. */
  getDisplayText(sheet: number, row: number, col: number): string | null {
    assertAlive(this);
    if (!this.capabilities.displayText) return null;
    return orFallback(null, () => ws(this, sheet).getDisplayText(row, col).text);
  }

  /** Format a UI value with a native Excel format code. This path is used
   *  when a store-side number format has not yet been flushed to the engine,
   *  while `getDisplayText` covers inherited and loaded formats. */
  formatValueForDisplay(value: CellValue, formatCode: string): string | null {
    assertAlive(this);
    if (!this.capabilities.displayText) return null;
    return orFallback(
      null,
      () => wb(this).formatValue(engineValueFromCell(value), formatCode).text,
    );
  }

  /** Native-name alias for callers that already have a format code. */
  formatValue(value: CellValue, formatCode: string): string | null {
    return this.formatValueForDisplay(value, formatCode);
  }

  /** State font 0, the record every cell with no font of its own resolves to.
   *  Fields the caller leaves out keep what the current default carries, except
   *  that a new name without a stated `scheme` drops the theme link, or the
   *  theme's font would be drawn instead of the named one. Returns false when
   *  the engine cannot restate its default. */
  setWorkbookDefaultFont(record: Partial<FontRecord>): boolean {
    assertAlive(this);
    if (!this.capabilities.defaultFont || !this.capabilities.cellFormatting) return false;
    const current = this.getFontRecord(0);
    if (!current) return false;
    const renamed =
      record.name !== undefined && record.name !== current.name && record.scheme === undefined;
    return orFallback(false, () => {
      wb(this).setDefaultFont(
        completeFontRecord({
          ...current,
          ...record,
          ...(renamed ? { scheme: FontScheme.None } : {}),
        }),
      );
      return true;
    });
  }

  /** Add or dedup a font record. Returns the resolved font index, or -1 on
   *  engine failure or when capability is off. */
  addFontRecord(record: FontRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    return orFallback(-1, () => wb(this).ooxml.addFont(completeFontRecord(record)));
  }

  /** Add or dedup a fill record. */
  addFillRecord(record: FillRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    return orFallback(-1, () => wb(this).ooxml.addFill(completeFillRecord(record)));
  }

  /** Add or dedup a border record. */
  addBorderRecord(record: BorderRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    return orFallback(-1, () => wb(this).ooxml.addBorder(completeBorderRecord(record)));
  }

  /** Register a number-format code. Built-in matches return the built-in id;
   *  custom codes are appended starting at 164. Returns -1 on failure. */
  addNumFmtCode(formatCode: string): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    return orFallback(-1, () => wb(this).ooxml.addNumberFormat(formatCode));
  }

  /** Add or dedup an XF (eXtended Format) record built from existing
   *  font/fill/border indices and a registered numFmtId. Returns the resolved
   *  xf index, or -1 on failure. */
  addXfRecord(record: CellXfInput): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    return orFallback(-1, () => wb(this).ooxml.addCellXf(record));
  }

  /** Number of `<cellStyle>` entries (named styles) registered on the
   *  workbook. Returns `0` under stub mode. */
  cellStyleCount(): number {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return 0;
    return orFallback(0, () => wb(this).listCellStyles().length);
  }

  /** Number of `<cellStyleXfs>` records — the named-style xf table that
   *  `getCellStyle(index).xfId` indexes into. Returns `0` under stub mode. */
  cellStyleXfCount(): number {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return 0;
    return orFallback(0, () => wb(this).ooxml.listCellStyleXfs().length);
  }

  /** Snapshot of the named cell style at `index`. Returns `null` when the
   *  index is out of range. */
  getCellStyle(index: number): {
    name: string;
    xfId: number;
    builtinId: number;
    iLevel: number;
    hidden: boolean;
    customBuiltin: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return null;
    const r = orFallback(null, () => wb(this).listCellStyles()[index] ?? null);
    if (!r) return null;
    return {
      name: r.name,
      xfId: r.styleXfIndex,
      builtinId: r.builtinId ?? CUSTOM_CELL_STYLE_BUILTIN_ID,
      iLevel: r.outlineLevel,
      hidden: r.hidden,
      customBuiltin: r.customBuiltin,
    };
  }

  /** Resolve the `<cellStyleXfs>` row used by a named style. This is
   *  separate from cell XF rows: `getCellStyle(index).xfId` points here. */
  getCellStyleXf(xfId: number): CellXf | null {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return null;
    return orFallback(null, () => wb(this).ooxml.getCellStyleXf(xfId));
  }

  /** Append (deduplicating) a `<cellStyleXfs>` row and return its index, which
   *  is what `setNamedCellStyle` and a cell XF's `styleXfIndex` reference.
   *  Returns -1 when the engine cannot author named styles. */
  addCellStyleXfRecord(record: CellXfInput): number {
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return -1;
    return orFallback(-1, () => wb(this).ooxml.addCellStyleXf(record));
  }

  /** Add or replace the `<cellStyle>` entry named `name`. `builtinId` is an
   *  OOXML ordinal, or null for a style with no built-in counterpart. */
  setNamedCellStyle(name: string, xfId: number, builtinId: number | null): boolean {
    return this.setNamedCellStyleRecord({ name, xfId, builtinId });
  }

  /** Full-record variant used when a rollback must preserve metadata that the
   *  three-argument public convenience method does not expose. */
  setNamedCellStyleRecord(record: NamedCellStyleRecord): boolean {
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return false;
    const builtinId =
      record.builtinId === null || record.builtinId === CUSTOM_CELL_STYLE_BUILTIN_ID
        ? null
        : record.builtinId;
    return orFallback(false, () => {
      wb(this).setCellStyle({
        name: record.name,
        styleXfIndex: record.xfId,
        builtinId,
        ...(record.iLevel === undefined ? {} : { outlineLevel: record.iLevel }),
        ...(record.hidden === undefined ? {} : { hidden: record.hidden }),
        ...(record.customBuiltin === undefined ? {} : { customBuiltin: record.customBuiltin }),
      });
      return true;
    });
  }

  /** Remove a named style. */
  removeNamedCellStyle(name: string): boolean {
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return false;
    return orFallback(false, () => {
      wb(this).removeCellStyle(name);
      return true;
    });
  }

  /** Enumerate every named cell style on the workbook — combines
   *  `cellStyleCount` + `getCellStyle` into one snapshot suitable for
   *  populating a "Cell Styles" UI. Empty under stub mode. Hidden
   *  built-ins are filtered out — the gallery hides those by default. */
  getNamedCellStyles(): readonly NamedCellStyle[] {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return [];
    return computeNamedCellStyles(this as unknown as WorkbookHandle);
  }
}

/** The stored name of the named cell style `name` resolves to under Excel's
 *  name identity; null when the workbook has none. */
export function storedCellStyleName(handle: WorkbookHandle, name: string): string | null {
  if (!handle.capabilities.cellStyles) return null;
  return orFallback(null, () => wb(handle).getCellStyle(name)?.name ?? null);
}

export function installStylesMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleStylesMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleStylesMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
