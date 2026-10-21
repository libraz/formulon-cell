import type {
  DisplayTextResult,
  EffectiveStyleResult,
  ThemeFonts,
  ThemeResult,
  Value,
} from '@libraz/formulon';
import { computeNamedCellStyles, type NamedCellStyle } from './cell-styles-meta.js';
import { completeBorderRecord, completeFillRecord, completeFontRecord } from './style-records.js';
import type {
  BorderRecord,
  CellValue,
  CellXf,
  ColorSpec,
  EngineCapabilities,
  FillRecord,
  FontRecord,
  Range,
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

/** `setCellStyle`'s sentinel for "this style has no OOXML built-in id". */
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

/** Theme read per handle; the engine theme only changes through the setters
 *  below, and painting resolves a theme-linked font for every visible cell. */
const themeCache = new WeakMap<object, ThemeResult | null>();

function readTheme(handle: WorkbookHandleStylesMethods): ThemeResult | null {
  if (themeCache.has(handle)) return themeCache.get(handle) ?? null;
  const native = wb(handle) as unknown as {
    getTheme?: () => ThemeResult;
  };
  const result = typeof native.getTheme === 'function' ? native.getTheme() : null;
  const theme = result?.status.ok ? result : null;
  themeCache.set(handle, theme);
  return theme;
}

function resolveArgb(handle: unknown, spec: ColorSpec, context: number): number | null {
  const capabilities = (handle as { capabilities?: EngineCapabilities }).capabilities;
  if (spec.kind === 0 || !capabilities?.colorResolution) return null;
  const native = wb(handle) as unknown as {
    resolveColor?: (
      color: ColorSpec,
      colorContext: number,
    ) => { status: { ok: boolean }; argb: number };
  };
  if (typeof native.resolveColor !== 'function') return null;
  const resolved = native.resolveColor(spec, context);
  return resolved.status.ok ? resolved.argb : null;
}

const engineValueFromCell = (value: CellValue): Value => {
  switch (value.kind) {
    case 'number':
      return { kind: 1, number: value.value, boolean: 0, text: '', errorCode: 0 };
    case 'bool':
      return { kind: 2, number: 0, boolean: value.value ? 1 : 0, text: '', errorCode: 0 };
    case 'text':
      return { kind: 3, number: 0, boolean: 0, text: value.value, errorCode: 0 };
    case 'error':
      return { kind: 4, number: 0, boolean: 0, text: '', errorCode: value.code };
    default:
      return { kind: 0, number: 0, boolean: 0, text: '', errorCode: 0 };
  }
};

export abstract class WorkbookHandleStylesMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Set the native cell-style XF for an inclusive column range. Width,
   * hidden and outline metadata stay owned by the engine's column record. */
  setColumnStyle(sheet: number, first: number, last: number, xfIndex: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    const native = wb(this) as unknown as {
      setColumnStyle?: (
        sheetIndex: number,
        firstColumn: number,
        lastColumn: number,
        styleXf: number,
      ) => { ok: boolean };
    };
    if (typeof native.setColumnStyle !== 'function') return false;
    return native.setColumnStyle(sheet, first, last, xfIndex).ok;
  }

  /** Clear the native cell-style XF for an inclusive column range. */
  clearColumnStyle(sheet: number, first: number, last: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    const native = wb(this) as unknown as {
      clearColumnStyle?: (
        sheetIndex: number,
        firstColumn: number,
        lastColumn: number,
      ) => { ok: boolean };
    };
    if (typeof native.clearColumnStyle !== 'function') return false;
    return native.clearColumnStyle(sheet, first, last).ok;
  }

  /** Set the native cell-style XF for one worksheet row. */
  setRowStyle(sheet: number, row: number, xfIndex: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    const native = wb(this) as unknown as {
      setRowStyle?: (...args: number[]) => { ok: boolean };
    };
    if (typeof native.setRowStyle !== 'function') return false;
    return (
      this.capabilities.rowStyleRanges
        ? native.setRowStyle(sheet, row, row, xfIndex)
        : native.setRowStyle(sheet, row, xfIndex)
    ).ok;
  }

  /** Clear the native cell-style XF for one worksheet row. */
  clearRowStyle(sheet: number, row: number): boolean {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return false;
    const native = wb(this) as unknown as {
      clearRowStyle?: (...args: number[]) => { ok: boolean };
    };
    if (typeof native.clearRowStyle !== 'function') return false;
    return (
      this.capabilities.rowStyleRanges
        ? native.clearRowStyle(sheet, row, row)
        : native.clearRowStyle(sheet, row)
    ).ok;
  }

  setRowStyleRange(sheet: number, first: number, last: number, xfIndex: number): boolean {
    assertAlive(this);
    if (!this.capabilities.rowStyleRanges) return false;
    const native = wb(this) as unknown as {
      setRowStyle: (...args: number[]) => { ok: boolean };
    };
    return native.setRowStyle(sheet, first, last, xfIndex).ok;
  }

  clearRowStyleRange(sheet: number, first: number, last: number): boolean {
    assertAlive(this);
    if (!this.capabilities.rowStyleRanges) return false;
    const native = wb(this) as unknown as {
      clearRowStyle: (...args: number[]) => { ok: boolean };
    };
    return native.clearRowStyle(sheet, first, last).ok;
  }

  getRangeStyleBases(range: Range): readonly number[] | null {
    assertAlive(this);
    if (!this.capabilities.rangeStylePatch) return null;
    const native = wb(this) as unknown as {
      rangeStyleBases: (...args: number[]) => { status: { ok: boolean }; xfIndices: number[] };
    };
    const result = native.rangeStyleBases(range.sheet, range.r0, range.c0, range.r1, range.c1);
    return result.status.ok ? [...result.xfIndices] : null;
  }

  remapRangeStyle(range: Range, fromXfs: readonly number[], toXfs: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.rangeStylePatch) return false;
    const native = wb(this) as unknown as {
      remapRangeStyle: (
        sheet: number,
        r0: number,
        c0: number,
        r1: number,
        c1: number,
        fromXfs: number[],
        toXfs: number[],
      ) => { ok: boolean };
    };
    return native.remapRangeStyle(
      range.sheet,
      range.r0,
      range.c0,
      range.r1,
      range.c1,
      [...fromXfs],
      [...toXfs],
    ).ok;
  }

  snapshotRangeStyle(range: Range): Uint8Array | null {
    assertAlive(this);
    if (!this.capabilities.rangeStyleSnapshot) return null;
    const native = wb(this) as unknown as {
      snapshotRangeStyle: (...args: number[]) => { status: { ok: boolean }; bytes: Uint8Array };
    };
    const result = native.snapshotRangeStyle(range.sheet, range.r0, range.c0, range.r1, range.c1);
    return result.status.ok ? new Uint8Array(result.bytes) : null;
  }

  restoreRangeStyle(sheet: number, bytes: Uint8Array): boolean {
    assertAlive(this);
    if (!this.capabilities.rangeStyleSnapshot) return false;
    const native = wb(this) as unknown as {
      restoreRangeStyle: (sheet: number, bytes: Uint8Array) => { ok: boolean };
    };
    return native.restoreRangeStyle(sheet, bytes).ok;
  }

  /** Capture only style-bearing axis records. Restoring these after clearing
   * an edited axis leaves width, hidden, outline, height and customHeight in
   * the native layout untouched. */
  getSheetFontAxisSnapshot(sheet: number): SheetFontAxisSnapshot | null {
    assertAlive(this);
    if (this.capabilities.sheetFontStyles !== true) return null;
    const native = wb(this) as unknown as {
      getSheetColumns?: (sheetIndex: number) => {
        status: { ok: boolean };
        columns?: ReadonlyArray<{
          first: number;
          last: number;
          hasStyle?: number | boolean;
          styleXf?: number;
        }>;
      };
      getSheetRowOverrides?: (sheetIndex: number) => {
        status: { ok: boolean };
        rows?: ReadonlyArray<{
          row: number;
          hasStyle?: number | boolean;
          styleXf?: number;
        }>;
      };
    };
    if (
      typeof native.getSheetColumns !== 'function' ||
      typeof native.getSheetRowOverrides !== 'function'
    )
      return null;
    const columnsResult = native.getSheetColumns(sheet);
    const rowsResult = native.getSheetRowOverrides(sheet);
    if (!columnsResult.status.ok || !rowsResult.status.ok) return null;
    const isSet = (value: number | boolean | undefined): boolean =>
      typeof value === 'boolean' ? value : value !== undefined && value !== 0;
    const columns = (columnsResult.columns ?? [])
      .filter((entry) => isSet(entry.hasStyle) && Number.isInteger(entry.styleXf))
      .map((entry) => ({
        first: entry.first,
        last: entry.last,
        styleXf: entry.styleXf as number,
      }));
    const rows = (rowsResult.rows ?? [])
      .filter((entry) => isSet(entry.hasStyle) && Number.isInteger(entry.styleXf))
      .map((entry) => ({ row: entry.row, styleXf: entry.styleXf as number }));
    return { columns, rows };
  }

  /** Read the XF (eXtended Format) table index assigned to `(sheet, row, col)`.
   *  Returns 0 (the workbook's default XF row) on missing cells. Returns null
   *  when the engine doesn't expose `getCellXfIndex` — i.e. the stub or older
   *  bundles. */
  getCellXfIndex(sheet: number, row: number, col: number): number | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getCellXfIndex(sheet, row, col);
    if (!r.status.ok) return null;
    return r.xfIndex;
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
    const s = wb(this).setCellXfIndex(sheet, row, col, xfIndex);
    return s.ok;
  }

  /** Remove a direct XF marker. Blank cells inherit their row/column style;
   * nonblank cells retain the native Cell/Normal resolution contract.
   * Values, formulas and annotations remain intact. Older bundles return
   * false rather than replacing the override with the explicit Normal XF. */
  clearCellXfIndex(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    const native = wb(this) as unknown as {
      clearCellXfIndex?: (sheet: number, row: number, col: number) => { ok: boolean };
    };
    if (typeof native.clearCellXfIndex !== 'function') return false;
    return native.clearCellXfIndex(sheet, row, col).ok;
  }

  /** Resolve the XF record at `xfIndex` to its component table indices
   *  (font / fill / border / number-format) plus alignment + wrap flags.
   *  Resolve those component indices with `getFontRecord`, `getFillRecord`,
   *  `getBorderRecord`, and `getNumFmtCode`. Returns null on engine failure
   *  or capability off. */
  getCellXf(xfIndex: number): CellXf | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getCellXf(xfIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Resolve a font index to its plain-data record. Returns null on engine
   *  failure or when `capabilities.cellFormatting` is off. */
  getFontRecord(fontIndex: number): FontRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getFont(fontIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Resolve a font record's theme, indexed, or automatic colour for display
   *  while retaining its raw selector. Theme-linked font names are projected
   *  from the current workbook theme for rendering; the returned `scheme`
   *  still preserves the source link so writeback can keep it linked. */
  getResolvedFontRecord(fontIndex: number): FontRecord | null {
    const record = this.getFontRecord(fontIndex);
    if (!record) return null;
    let resolved = record;
    if (record.scheme === 1 || record.scheme === 2) {
      const theme = this.capabilities.workbookTheme ? readTheme(this) : null;
      const eastAsian = record.hasCharset && record.charset !== 0;
      const themeName =
        record.scheme === 1
          ? eastAsian
            ? theme?.fonts.majorEastAsian
            : theme?.fonts.majorLatin
          : eastAsian
            ? theme?.fonts.minorEastAsian
            : theme?.fonts.minorLatin;
      if (themeName) resolved = { ...resolved, name: themeName };
    }
    if (!record.color) return resolved;
    const colorArgb = resolveArgb(this, record.color, 0);
    return colorArgb === null ? resolved : { ...resolved, colorArgb };
  }

  /** Resolve a fill index to its plain-data record. */
  getFillRecord(fillIndex: number): FillRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getFill(fillIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Resolve a fill record's theme, indexed, or automatic colours for display
   *  while retaining both raw selectors. RGB records are returned unchanged. */
  getResolvedFillRecord(fillIndex: number): FillRecord | null {
    const record = this.getFillRecord(fillIndex);
    if (!record) return null;
    const fgArgb = record.fg ? resolveArgb(this, record.fg, 1) : null;
    const bgArgb = record.bg ? resolveArgb(this, record.bg, 2) : null;
    if (fgArgb === null && bgArgb === null) return record;
    return {
      ...record,
      ...(fgArgb === null ? {} : { fgArgb }),
      ...(bgArgb === null ? {} : { bgArgb }),
    };
  }

  /** Resolve a border index to its plain-data record. */
  getBorderRecord(borderIndex: number): BorderRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getBorder(borderIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
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
        const colorArgb = current.color ? resolveArgb(this, current.color, 3) : null;
        if (colorArgb === null) {
          out[side] = current;
        } else {
          changed = true;
          out[side] = { ...current, colorArgb };
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
    const r = wb(this).getNumFmt(numFmtId);
    if (!r.status.ok) return null;
    return r.formatCode;
  }

  /** Read the workbook theme without exposing the raw WASM workbook. The
   *  returned selectors stay in clrScheme order so callers can edit a single
   *  preset and send the same twelve-entry array back to the engine. */
  getWorkbookTheme(): ThemeResult | null {
    assertAlive(this);
    if (!this.capabilities.workbookTheme) return null;
    const result = readTheme(this);
    if (!result) return null;
    return {
      ...result,
      status: { ...result.status },
      colors: [...result.colors],
      fonts: { ...result.fonts },
    };
  }

  /** Native-name alias for consumers that work directly with Workbook APIs. */
  getTheme(): ThemeResult | null {
    return this.getWorkbookTheme();
  }

  /** Replace the workbook's twelve clrScheme colours. */
  setWorkbookThemeColors(colors: readonly number[]): boolean {
    assertAlive(this);
    if (!this.capabilities.workbookTheme || colors.length !== 12) return false;
    const native = wb(this) as unknown as {
      setThemeColors?: (next: number[]) => { ok: boolean };
    };
    if (typeof native.setThemeColors !== 'function') return false;
    themeCache.delete(this);
    return native.setThemeColors([...colors]).ok;
  }

  /** Native-name alias for workbook theme editing. */
  setThemeColors(colors: readonly number[]): boolean {
    return this.setWorkbookThemeColors(colors);
  }

  /** Replace the workbook's major/minor Latin and East Asian theme faces. */
  setWorkbookThemeFonts(fonts: ThemeFonts): boolean {
    assertAlive(this);
    if (!this.capabilities.workbookTheme) return false;
    const native = wb(this) as unknown as {
      setThemeFonts?: (next: ThemeFonts) => { ok: boolean };
    };
    if (typeof native.setThemeFonts !== 'function') return false;
    themeCache.delete(this);
    return native.setThemeFonts({ ...fonts }).ok;
  }

  /** Native-name alias for workbook theme editing. */
  setThemeFonts(fonts: ThemeFonts): boolean {
    return this.setWorkbookThemeFonts(fonts);
  }

  /** Return the XF selected by cell, row, column, or workbook default. This
   *  remains a read-only projection: callers must not feed its XF index back
   *  to `setCellXfIndex` just because a row/column supplied the formatting. */
  getEffectiveStyle(sheet: number, row: number, col: number): EffectiveStyleResult | null {
    assertAlive(this);
    if (!this.capabilities.effectiveStyle) return null;
    const native = wb(this) as unknown as {
      getEffectiveStyle?: (
        sheetIndex: number,
        rowIndex: number,
        colIndex: number,
      ) => EffectiveStyleResult | undefined;
    };
    if (typeof native.getEffectiveStyle !== 'function') return null;
    const result = native.getEffectiveStyle(sheet, row, col);
    if (!result?.status.ok) return null;
    return {
      ...result,
      status: { ...result.status },
      font: { ...result.font },
      fillForeground: { ...result.fillForeground },
      fillBackground: { ...result.fillBackground },
      borders: {
        left: { ...result.borders.left },
        right: { ...result.borders.right },
        top: { ...result.borders.top },
        bottom: { ...result.borders.bottom },
        diagonal: { ...result.borders.diagonal },
      },
    };
  }

  /** Use formulon's workbook-aware display renderer for an existing cell. */
  getDisplayText(sheet: number, row: number, col: number): string | null {
    assertAlive(this);
    if (!this.capabilities.displayText) return null;
    const native = wb(this) as unknown as {
      getDisplayText?: (
        sheetIndex: number,
        rowIndex: number,
        colIndex: number,
      ) => DisplayTextResult;
    };
    if (typeof native.getDisplayText !== 'function') return null;
    const result = native.getDisplayText(sheet, row, col);
    return result.status.ok ? result.text : null;
  }

  /** Format a UI value with a native Excel format code. This path is used
   *  when a store-side number format has not yet been flushed to the engine,
   *  while `getDisplayText` covers inherited and loaded formats. */
  formatValueForDisplay(value: CellValue, formatCode: string): string | null {
    assertAlive(this);
    if (!this.capabilities.displayText) return null;
    const native = wb(this) as unknown as {
      formatValue?: (next: Value, code: string) => DisplayTextResult;
    };
    if (typeof native.formatValue !== 'function') return null;
    const result = native.formatValue(engineValueFromCell(value), formatCode);
    return result.status.ok ? result.text : null;
  }

  /** Native-name alias for callers that already have a format code. */
  formatValue(value: CellValue, formatCode: string): string | null {
    return this.formatValueForDisplay(value, formatCode);
  }

  /** State font 0, the record every cell with no font of its own resolves to.
   *  Fields the caller leaves out keep whatever the current default carries —
   *  notably its `<scheme>` theme link, which is what keeps the font tracking
   *  the workbook theme instead of becoming a literal name. Returns false when
   *  the engine cannot restate its default. */
  setWorkbookDefaultFont(record: Partial<FontRecord>): boolean {
    assertAlive(this);
    if (!this.capabilities.defaultFont || !this.capabilities.cellFormatting) return false;
    const current = this.getFontRecord(0);
    if (!current) return false;
    return wb(this).setDefaultFont(completeFontRecord({ ...current, ...record })).ok;
  }

  /** Add or dedup a font record. Returns the resolved font index, or -1 on
   *  engine failure or when capability is off. */
  addFontRecord(record: FontRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addFont(completeFontRecord(record));
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a fill record. */
  addFillRecord(record: FillRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addFill(completeFillRecord(record));
    return r.status.ok ? r.index : -1;
  }

  /** Add or dedup a border record. */
  addBorderRecord(record: BorderRecord): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addBorder(completeBorderRecord(record));
    return r.status.ok ? r.index : -1;
  }

  /** Register a number-format code. Built-in matches return the built-in id;
   *  custom codes are appended starting at 164. Returns -1 on failure. */
  addNumFmtCode(formatCode: string): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addNumFmt(formatCode);
    return r.status.ok ? r.numFmtId : -1;
  }

  /** Add or dedup an XF (eXtended Format) record built from existing
   *  font/fill/border indices and a registered numFmtId. Returns the resolved
   *  xf index, or -1 on failure. */
  addXfRecord(record: CellXf): number {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return -1;
    const r = wb(this).addXf(record);
    return r.status.ok ? r.index : -1;
  }

  /** Number of `<cellStyle>` entries (named styles) registered on the
   *  workbook. Returns `0` under stub mode and older engine package builds. */
  cellStyleCount(): number {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return 0;
    const count = wb(this).cellStyleCount();
    return count.status.ok ? count.value : 0;
  }

  /** Number of `<cellStyleXfs>` records — the named-style xf table that
   *  `CellStyleResult.xfId` indexes into. Returns `0` under stub mode and
   *  older engine package builds. */
  cellStyleXfCount(): number {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return 0;
    const count = wb(this).cellStyleXfCount();
    return count.status.ok ? count.value : 0;
  }

  /** Snapshot of the named cell style at `index`. Returns `null` when the
   *  engine doesn't expose `getCellStyle` or the index is out of range. */
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
    const r = wb(this).getCellStyle(index);
    if (!r.status.ok) return null;
    return {
      name: r.name,
      xfId: r.xfId,
      builtinId: r.builtinId,
      iLevel: r.iLevel,
      hidden: r.hidden,
      customBuiltin: r.customBuiltin,
    };
  }

  /** Resolve the `<cellStyleXfs>` row used by a named style. This is
   *  separate from cell XF rows: `getCellStyle(index).xfId` points here. */
  getCellStyleXf(xfId: number): CellXf | null {
    assertAlive(this);
    if (!this.capabilities.cellStyles) return null;
    const r = wb(this).getCellStyleXf(xfId);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  /** Append (deduplicating) a `<cellStyleXfs>` row and return its index, which
   *  is what `setNamedCellStyle` and a cell XF's `xfId` reference. Returns -1
   *  when the engine cannot author named styles. */
  addCellStyleXfRecord(record: CellXf): number {
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return -1;
    const r = wb(this).addCellStyleXf(record);
    return r.status.ok ? r.index : -1;
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
    return wb(this).setCellStyle({
      name: record.name,
      xfId: record.xfId,
      builtinId: record.builtinId ?? CUSTOM_CELL_STYLE_BUILTIN_ID,
      ...(record.iLevel === undefined ? {} : { iLevel: record.iLevel }),
      ...(record.hidden === undefined ? {} : { hidden: record.hidden }),
      ...(record.customBuiltin === undefined ? {} : { customBuiltin: record.customBuiltin }),
    }).ok;
  }

  /** Remove a named style when the loaded engine exposes the 0.13 API. */
  removeNamedCellStyle(name: string): boolean {
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return false;
    const native = wb(this) as unknown as {
      removeCellStyle?: (styleName: string) => { ok: boolean };
    };
    if (typeof native.removeCellStyle !== 'function') return false;
    return native.removeCellStyle(name).ok;
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

export function installStylesMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleStylesMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleStylesMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
