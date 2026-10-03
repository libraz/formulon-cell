import { computeNamedCellStyles, type NamedCellStyle } from './cell-styles-meta.js';
import { completeBorderRecord, completeFillRecord, completeFontRecord } from './style-records.js';
import type {
  BorderRecord,
  CellXf,
  EngineCapabilities,
  FillRecord,
  FontRecord,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

/** `setCellStyle`'s sentinel for "this style has no OOXML built-in id". The
 *  binding takes a fixed argument count, so the absence has to be a value. */
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

export abstract class WorkbookHandleStylesMethods {
  declare readonly capabilities: EngineCapabilities;

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
   *  point at an existing row in the workbook's XF table — there is no
   *  upstream API to insert new XF rows yet, so this is mainly useful for
   *  cloning formatting from one cell to another (Format-Painter parity at
   *  the engine layer) or for clearing back to xfIndex 0 (the default).
   *  Returns false on engine failure or when `capabilities.cellFormatting`
   *  is off. */
  setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): boolean {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return false;
    const s = wb(this).setCellXfIndex(sheet, row, col, xfIndex);
    return s.ok;
  }

  /** Resolve the XF record at `xfIndex` to its component table indices
   *  (font / fill / border / number-format) plus alignment + wrap flags.
   *  Note that the component indices are themselves opaque without
   *  resolver APIs (`getFont(idx)`, `getFill(idx)`, …) which upstream has
   *  not exposed yet — so this is currently most useful as a metadata
   *  signal (e.g. "do these two cells share the same XF row?"). Returns
   *  null on engine failure or capability off. */
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

  /** Resolve a fill index to its plain-data record. */
  getFillRecord(fillIndex: number): FillRecord | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getFill(fillIndex);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
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

  /** Resolve a number-format id to its format-code string. */
  getNumFmtCode(numFmtId: number): string | null {
    assertAlive(this);
    if (!this.capabilities.cellFormatting) return null;
    const r = wb(this).getNumFmt(numFmtId);
    if (!r.status.ok) return null;
    return r.formatCode;
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
    assertAlive(this);
    if (!this.capabilities.cellStyleMutate) return false;
    return wb(this).setCellStyle(name, xfId, builtinId ?? CUSTOM_CELL_STYLE_BUILTIN_ID).ok;
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
