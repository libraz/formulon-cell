import type {
  Addr,
  CellValue,
  EngineCapabilities,
  FormulonModule,
  FunctionMetadataProvider,
  PhoneticRun,
  TableInput,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  module: FormulonModule;
  capabilities: EngineCapabilities;
  functionMetadataProvider: FunctionMetadataProvider | null;
  assertAlive(): void;
};
type EngineCommentEntry = { row: number; col: number; author: string; text: string };
type CommentEnumerableWorkbook = Workbook & {
  getComments?: (
    sheet: number,
  ) => readonly EngineCommentEntry[] & { readonly status: { ok: boolean } };
};
type TableAuthoringWorkbook = Workbook & {
  createTable?: (input: TableInput) => { status: { ok: boolean }; index: number };
  updateTable?: (
    index: number,
    input: Pick<TableInput, 'ref' | 'styleName' | 'headerRow' | 'totalsRow'>,
  ) => { ok: boolean };
  removeTable?: (index: number) => { ok: boolean };
};
type AutoFilterWorkbook = Workbook & {
  getSheetAutoFilterXml?: (sheet: number) => { status: { ok: boolean }; xml: string };
  setSheetAutoFilterXml?: (sheet: number, xml: string) => { ok: boolean };
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleFeatureMethods {}
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

export abstract class WorkbookHandleFeatureMethods {
  declare readonly capabilities: EngineCapabilities;
  declare readonly sheetCount: number;
  abstract getValue(addr: Addr): CellValue;

  /** Returns the complete worksheet `<autoFilter>` fragment, or null when
   * the engine does not expose the optional raw-definition seam. */
  getSheetAutoFilterXml(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return null;
    const r = (wb(this) as AutoFilterWorkbook).getSheetAutoFilterXml?.(sheet);
    return r?.status.ok ? r.xml : null;
  }

  /** Replaces the worksheet `<autoFilter>` fragment. Empty XML removes the
   * definition. Returns false for old bundles and failed writes. */
  setSheetAutoFilterXml(sheet: number, xml: string): boolean {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return false;
    return (wb(this) as AutoFilterWorkbook).setSheetAutoFilterXml?.(sheet, xml)?.ok === true;
  }

  /** Read the cell's OOXML phonetic guide, if the current engine exposes it. */
  getCellPhonetic(sheet: number, row: number, col: number): string | null {
    assertAlive(this);
    if (!this.capabilities.phonetic) return null;
    const r = wb(this).getCellPhonetic(sheet, row, col);
    return r.status.ok && r.value ? r.value : null;
  }

  /** Set (or, with an empty string, clear) the cell's phonetic guide. The
   *  engine spans the whole cell text, so this replaces any per-run guide the
   *  cell carried — use `setCellPhoneticRuns` to preserve the spans. */
  setCellPhonetic(sheet: number, row: number, col: number, phonetic: string): boolean {
    assertAlive(this);
    if (!this.capabilities.phonetic) return false;
    return wb(this).setCellPhonetic(sheet, row, col, phonetic).ok;
  }

  /** Read the cell's phonetic guide span by span. Returns null when the engine
   *  has no per-run surface, which is distinct from the empty array an
   *  unannotated cell reports. */
  getCellPhoneticRuns(sheet: number, row: number, col: number): PhoneticRun[] | null {
    assertAlive(this);
    if (!this.capabilities.phoneticRuns) return null;
    const r = wb(this).getCellPhoneticRuns(sheet, row, col);
    if (!r.status.ok) return null;
    return r.runs.map((run) => ({ start: run.sb, end: run.eb, text: run.text }));
  }

  /** Replace the cell's phonetic guide with `runs`, an ordered partition of the
   *  cell text. An empty array clears the guide. */
  setCellPhoneticRuns(
    sheet: number,
    row: number,
    col: number,
    runs: readonly PhoneticRun[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.phoneticRuns) return false;
    return wb(this).setCellPhoneticRuns(
      sheet,
      row,
      col,
      runs.map((run) => ({ sb: run.start, eb: run.end, text: run.text })),
    ).ok;
  }

  /** Creates an OOXML worksheet table and returns its index, or -1 when the
   * loaded engine predates table authoring support. */
  createTable(input: TableInput): number {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return -1;
    const result = (wb(this) as TableAuthoringWorkbook).createTable?.(input);
    return result?.status.ok ? result.index : -1;
  }

  updateTable(
    index: number,
    input: Pick<TableInput, 'ref' | 'styleName' | 'headerRow' | 'totalsRow'>,
  ): boolean {
    assertAlive(this);
    return (
      this.capabilities.tableMutate === true &&
      (wb(this) as TableAuthoringWorkbook).updateTable?.(index, input).ok === true
    );
  }

  removeTable(index: number): boolean {
    assertAlive(this);
    return (
      this.capabilities.tableMutate === true &&
      (wb(this) as TableAuthoringWorkbook).removeTable?.(index).ok === true
    );
  }

  /** Read the cell comment at `(sheet, row, col)`. Returns null when the
   *  cell has no comment or when the engine doesn't expose `getComment`. */
  getComment(sheet: number, row: number, col: number): { author: string; text: string } | null {
    assertAlive(this);
    if (!this.capabilities.comments) return null;
    const e = wb(this).getComment(sheet, row, col);
    return e ? { author: e.author, text: e.text } : null;
  }

  /** Snapshot every comment on `sheet` when the engine exposes a sheet-wide
   *  enumerator. Empty under stub or older engines. */
  getComments(sheet: number): { row: number; col: number; author: string; text: string }[] {
    assertAlive(this);
    if (!this.capabilities.commentsEnumerable) return [];
    const engineWb = wb(this) as CommentEnumerableWorkbook;
    if (typeof engineWb.getComments !== 'function') return [];
    const entries = engineWb.getComments(sheet);
    if (!entries.status.ok) return [];
    return entries.map((e) => ({
      row: e.row,
      col: e.col,
      author: e.author,
      text: e.text,
    }));
  }

  /** Persist a cell comment. Empty `text` removes it. No-op (returns false)
   *  under the stub. */
  setCommentEntry(sheet: number, row: number, col: number, author: string, text: string): boolean {
    assertAlive(this);
    if (!this.capabilities.comments) return false;
    const s = wb(this).setComment(sheet, row, col, author, text);
    return s.ok;
  }

  /** Reads the round-trip `<sheetProtection>` flags. Returns `null` when
   *  the engine doesn't expose `getSheetProtection`. The booleans are
   *  reported as JS booleans (the engine wires them as 0/1 numbers); the
   *  `enabled` flag denotes whether the protection block is emitted on
   *  save. */
  getSheetProtection(sheet: number): {
    enabled: boolean;
    algorithmName: string;
    hashValue: string;
    saltValue: string;
    spinCount: number;
    legacyPassword: string;
    sheet: boolean;
    objects: boolean;
    scenarios: boolean;
    formatCells: boolean;
    formatColumns: boolean;
    formatRows: boolean;
    insertColumns: boolean;
    insertRows: boolean;
    insertHyperlinks: boolean;
    deleteColumns: boolean;
    deleteRows: boolean;
    selectLockedCells: boolean;
    selectUnlockedCells: boolean;
    sort: boolean;
    autoFilter: boolean;
    pivotTables: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetProtectionRoundtrip) return null;
    const r = wb(this).getSheetProtection(sheet);
    if (!r.status.ok) return null;
    const p = r.protection;
    return {
      enabled: p.enabled !== 0,
      algorithmName: p.algorithmName,
      hashValue: p.hashValue,
      saltValue: p.saltValue,
      spinCount: p.spinCount,
      legacyPassword: p.legacyPassword,
      sheet: p.sheet !== 0,
      objects: p.objects !== 0,
      scenarios: p.scenarios !== 0,
      formatCells: p.formatCells !== 0,
      formatColumns: p.formatColumns !== 0,
      formatRows: p.formatRows !== 0,
      insertColumns: p.insertColumns !== 0,
      insertRows: p.insertRows !== 0,
      insertHyperlinks: p.insertHyperlinks !== 0,
      deleteColumns: p.deleteColumns !== 0,
      deleteRows: p.deleteRows !== 0,
      selectLockedCells: p.selectLockedCells !== 0,
      selectUnlockedCells: p.selectUnlockedCells !== 0,
      sort: p.sort !== 0,
      autoFilter: p.autoFilter !== 0,
      pivotTables: p.pivotTables !== 0,
    };
  }

  /** Replaces `<sheetProtection>` flags wholesale. Setting `enabled` to
   *  `false` clears the protection block on save. Returns `false` (no-op)
   *  under stub mode and older engine package builds. */
  setSheetProtection(
    sheet: number,
    protection: {
      enabled: boolean;
      legacyPassword?: string;
      algorithmName?: string;
      hashValue?: string;
      saltValue?: string;
      spinCount?: number;
      sheet?: boolean;
      objects?: boolean;
      scenarios?: boolean;
      formatCells?: boolean;
      formatColumns?: boolean;
      formatRows?: boolean;
      insertColumns?: boolean;
      insertRows?: boolean;
      insertHyperlinks?: boolean;
      deleteColumns?: boolean;
      deleteRows?: boolean;
      selectLockedCells?: boolean;
      selectUnlockedCells?: boolean;
      sort?: boolean;
      autoFilter?: boolean;
      pivotTables?: boolean;
    },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetProtectionRoundtrip) return false;
    const b = (v: boolean | undefined): number => (v ? 1 : 0);
    const s = wb(this).setSheetProtection(sheet, {
      enabled: b(protection.enabled),
      algorithmName: protection.algorithmName ?? '',
      hashValue: protection.hashValue ?? '',
      saltValue: protection.saltValue ?? '',
      spinCount: protection.spinCount ?? 0,
      legacyPassword: protection.legacyPassword ?? '',
      sheet: b(protection.sheet ?? true),
      objects: b(protection.objects),
      scenarios: b(protection.scenarios),
      formatCells: b(protection.formatCells),
      formatColumns: b(protection.formatColumns),
      formatRows: b(protection.formatRows),
      insertColumns: b(protection.insertColumns),
      insertRows: b(protection.insertRows),
      insertHyperlinks: b(protection.insertHyperlinks),
      deleteColumns: b(protection.deleteColumns),
      deleteRows: b(protection.deleteRows),
      selectLockedCells: b(protection.selectLockedCells),
      selectUnlockedCells: b(protection.selectUnlockedCells),
      sort: b(protection.sort),
      autoFilter: b(protection.autoFilter),
      pivotTables: b(protection.pivotTables),
    });
    return s.ok;
  }
}

export function installWorkbookFeatureMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleFeatureMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleFeatureMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
